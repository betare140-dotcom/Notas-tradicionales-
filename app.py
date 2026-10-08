from hashlib import sha256
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
import streamlit as st
from extractor import FIELDS, SUMMARY_ALIASES, text, normalize, create_word, detect_header, display_date, load_excel, read_records, suggest_mapping
from resumenes import fetch_excerpt, make_excerpt

def main():
    st.set_page_config(page_title='Medios tradicionales', page_icon='📰', layout='wide')
    st.title('Extractor de medios tradicionales')
    st.write('Carga tu Excel y descarga las notas agrupadas por fecha, con medio, título, extracto y enlace activo.')
    uploaded = st.file_uploader('Archivo Excel', type=['xlsx'])
    if uploaded is None:
        st.info('Compatible con la hoja Noticias y sus columnas Portal, Fecha, Título descriptivo y URL. También puedes elegir otras columnas.')
        return
    payload = uploaded.getvalue()
    try:
        workbook = load_excel(payload)
    except Exception as exc:
        st.error(f'No se pudo abrir el Excel: {exc}')
        return
    fingerprint = sha256(payload).hexdigest()[:16]
    sheet_name = st.selectbox('Hoja que deseas extraer', workbook.sheetnames, key=f'sheet_{fingerprint}')
    sheet = workbook[sheet_name]
    if not sheet.max_row:
        st.warning('La hoja está vacía.'); return
    scope = f'{fingerprint}_{sheet_name}'
    header_row = st.number_input('Fila de encabezados', min_value=1, max_value=sheet.max_row,
                                 value=detect_header(sheet), step=1, key=f'header_{scope}')
    headers = [c.value for c in sheet[int(header_row)]]
    suggestions = suggest_mapping(headers)
    labels = ['Seleccionar columna'] + [f'{i+1} · {h if h is not None else "Sin encabezado"}' for i, h in enumerate(headers)]
    with st.expander('Columnas del reporte', expanded=True):
        mapping = {}
        for field, container in zip(FIELDS, st.columns(4)):
            suggested = suggestions[field]
            selected = container.selectbox(field, range(len(labels)),
                index=suggested+1 if suggested is not None else 0,
                format_func=lambda i: labels[i], key=f'map_{scope}_{header_row}_{field}')
            if selected: mapping[field] = selected - 1
        summary_default = next((i+1 for i, h in enumerate(headers) if normalize(h) in SUMMARY_ALIASES), 0)
        summary_choice = st.selectbox('Columna de extracto, resumen o contenido (opcional)', range(len(labels)),
            index=summary_default, format_func=lambda i: 'No hay texto en el Excel' if i == 0 else labels[i],
            key=f'summary_map_{scope}_{header_row}')
    if len(mapping) != 4:
        st.warning('Selecciona las cuatro columnas para continuar.'); return
    if len(set(mapping.values())) != 4:
        st.warning('Elige una columna distinta para cada campo.'); return
    records = read_records(sheet, int(header_row), mapping, workbook.epoch, summary_choice-1 if summary_choice else None)
    state_key = sha256(f'{scope}_{header_row}_{mapping}_{summary_choice}'.encode()).hexdigest()
    summaries = st.session_state.setdefault(f'summaries_{state_key}', {})
    for r in records:
        initial = r['Resumen']
        if summary_choice and normalize(headers[summary_choice-1]) in ('contenido', 'texto', 'contenido de la nota'):
            initial = make_excerpt(initial, r['Título'])
        summaries.setdefault(r['_row'], {'Resumen': initial, 'Estado': 'Texto del Excel' if initial else 'Pendiente', 'Incluir': True})
    if not records:
        st.warning('No se encontraron notas debajo de los encabezados.'); return
    flagged = [r for r in records if r['_warnings']]
    if flagged:
        st.warning(f'{len(flagged)} filas tienen datos faltantes o por revisar. Se conservarán salvo que las excluyas.')
        with st.expander('Revisar datos del archivo'):
            st.dataframe(pd.DataFrame([{'Fila Excel': r['_row'], 'Observaciones': '; '.join(r['_warnings'])} for r in flagged]), hide_index=True)
    media = sorted({r['Medio'] or 'Sin medio' for r in records})
    chosen = st.multiselect('Medios incluidos', media, default=media, key=f'media_{scope}')
    selected = [r for r in records if (r['Medio'] or 'Sin medio') in chosen]
    if st.checkbox('Filtrar por fechas', key=f'use_dates_{scope}'):
        dates = [r['_date'] for r in records if r['_date']]
        if dates:
            c1, c2 = st.columns(2)
            start = c1.date_input('Desde', min(dates), key=f'start_{scope}')
            end = c2.date_input('Hasta', max(dates), key=f'end_{scope}')
            keep_unknown = st.checkbox('Incluir filas con fecha no reconocida', value=True, key=f'unknown_{scope}')
            if start > end:
                st.error('La fecha inicial debe ser anterior o igual a la final.'); return
            selected = [r for r in selected if (start <= r['_date'] <= end if r['_date'] else keep_unknown)]
        else:
            st.warning('No hay fechas reconocidas; se conservan las filas.')
    if st.checkbox('Eliminar duplicados exactos de fecha, medio, título y link', value=False):
        seen = set(); unique = []
        for r in selected:
            key = (display_date(r), r['Medio'], r['Título'], r['Link'])
            if key not in seen: unique.append(r); seen.add(key)
        st.caption(f'Se quitaron {len(selected)-len(unique)} duplicados exactos.')
        selected = unique
    order = st.selectbox('Orden del Word', ['Fecha ascendente', 'Fecha descendente', 'Orden del Excel'])
    if order != 'Orden del Excel':
        dated = sorted([r for r in selected if r['_date']], key=lambda r: r['_date'], reverse=order == 'Fecha descendente')
        selected = dated + [r for r in selected if not r['_date']]
    st.subheader('Revisa las notas')
    st.caption('Desmarca Incluir para excluir una nota. Puedes escribir o corregir el extracto en la columna Resumen.')
    if not selected:
        st.info('No hay notas con los filtros seleccionados.'); return
    st.write('El extracto automático selecciona oraciones del texto de la noticia relacionadas con el título. Revísalo para confirmar el punto importante.')
    if st.button('Obtener extractos desde los enlaces', type='primary'):
        pending = [r for r in selected if summaries[r['_row']]['Incluir'] and not summaries[r['_row']]['Resumen'].strip()]
        if not pending:
            st.info('Las notas incluidas ya tienen un extracto. Puedes editarlo o vaciarlo para obtenerlo otra vez.')
        else:
            progress = st.progress(0, text='Leyendo noticias…')
            with ThreadPoolExecutor(max_workers=4) as pool:
                jobs = {pool.submit(fetch_excerpt, r['Link'], r['Título']): r for r in pending}
                for n, future in enumerate(as_completed(jobs), 1):
                    row = jobs[future]['_row']
                    try:
                        summaries[row].update(future.result())
                    except Exception as exc:
                        summaries[row]['Estado'] = str(exc)[:240]
                    progress.progress(n / len(pending), text=f'{n} de {len(pending)} noticias revisadas')
            st.session_state[f'revision_{state_key}'] = st.session_state.get(f'revision_{state_key}', 0) + 1
    preview = pd.DataFrame([{'Incluir': summaries[r['_row']]['Incluir'], 'Fila Excel': r['_row'], 'Fecha': display_date(r),
        'Medio': r['Medio'], 'Título': r['Título'], 'Resumen': summaries[r['_row']]['Resumen'],
        'Link': r['Link'], 'Estado del extracto': summaries[r['_row']]['Estado']} for r in selected])
    editor_key = sha256((state_key + str([r['_row'] for r in selected])).encode()).hexdigest()
    revision = st.session_state.get(f'revision_{state_key}', 0)
    edited = st.data_editor(preview, hide_index=True, width='stretch', disabled=['Fila Excel', *FIELDS, 'Estado del extracto'],
        column_config={'Incluir': st.column_config.CheckboxColumn('Incluir'), 'Link': st.column_config.LinkColumn('Link')},
        key=f'preview_{editor_key}_{revision}')
    included = []
    for r, (_, edited_row) in zip(selected, edited.iterrows()):
        summary = text(edited_row['Resumen'])
        if summary != summaries[r['_row']]['Resumen']:
            summaries[r['_row']]['Estado'] = 'Editado manualmente' if summary else 'Pendiente'
        summaries[r['_row']]['Resumen'] = summary
        summaries[r['_row']]['Incluir'] = bool(edited_row['Incluir'])
        if edited_row['Incluir']:
            included.append({**r, 'Resumen': summary})
    st.success(f'{len(included)} notas seleccionadas de {len(records)} filas leídas.')
    title = st.text_input('Título del documento', 'Notas de medios tradicionales')
    if included:
        missing = sum(not r['Resumen'] for r in included)
        if missing:
            st.warning(f'Faltan {missing} extractos. Completa la columna Resumen; si descargas ahora, aparecerán como pendientes.')
        data = create_word(included, title or 'Notas de medios tradicionales')
        st.download_button('Descargar Word', data, file_name='Notas_medios_tradicionales.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', type='primary')

if __name__ == '__main__':
    main()
