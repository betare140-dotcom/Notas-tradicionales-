"""Extractos de noticias sin API: oraciones del contenido, editables antes de exportar."""
from html.parser import HTMLParser
import ipaddress
import json
import re
import socket
from urllib.parse import urljoin, urlsplit
from urllib.request import Request, build_opener, HTTPRedirectHandler
from urllib.error import HTTPError
from extractor import normalize, text, valid_url

class NewsParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack = []; self.paragraph = None; self.paragraphs = []
        self.script = None; self.scripts = []; self.descriptions = []
        self.page_title = []; self.in_title = False

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == 'meta':
            if attrs.get('property', attrs.get('name', '')).lower() in ('og:description', 'description', 'twitter:description'):
                self.descriptions.append(attrs.get('content', ''))
        if tag in ('meta', 'link', 'img', 'br', 'input', 'hr', 'source', 'wbr', 'area', 'base', 'embed', 'param', 'track'):
            if tag == 'br' and self.paragraph is not None: self.paragraph[1].append(' ')
            return
        marker = ' '.join([attrs.get('class', ''), attrs.get('id', ''), attrs.get('itemprop', '')]).lower()
        self.stack.append((tag, marker))
        if tag == 'title': self.in_title = True
        if tag == 'script' and attrs.get('type', '').lower() == 'application/ld+json': self.script = []
        if tag == 'p':
            if self.paragraph is not None: self._finish_paragraph()
            contexts = ' '.join(m for t, m in self.stack)
            explicit_body = any(key in contexts for key in ('articlebody', 'article-body', 'article-content', 'entry-content', 'field--name-body', 'field-name-body', 'post-content', 'story-body', 'td-post-content', 'article__body', 'contenido-nota'))
            is_content = 2 if explicit_body else (1 if any(t == 'article' for t, m in self.stack) else 0)
            is_noise = any(t in ('nav', 'footer', 'aside') for t, m in self.stack) or any(
                key in contexts for key in ('related', 'comment', 'newsletter', 'author-bio', 'social-share'))
            self.paragraph = (is_content if not is_noise else 0, [])

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in ('meta', 'link', 'img', 'br', 'input', 'hr', 'source'): self.handle_endtag(tag)

    def handle_endtag(self, tag):
        if tag == 'p': self._finish_paragraph()
        if tag == 'script' and self.script is not None:
            self.scripts.append(''.join(self.script)); self.script = None
        if tag == 'title': self.in_title = False
        for i in range(len(self.stack)-1, -1, -1):
            if self.stack[i][0] == tag:
                del self.stack[i:]; break

    def handle_data(self, value):
        if self.script is not None: self.script.append(value)
        if self.in_title: self.page_title.append(value)
        if self.paragraph is not None and not any(t in ('script', 'style') for t, _ in self.stack):
            self.paragraph[1].append(value)

    def _finish_paragraph(self):
        if self.paragraph is not None:
            priority, parts = self.paragraph
            value = re.sub(r'\s+', ' ', ''.join(parts)).strip()
            if len(value.split()) >= 8: self.paragraphs.append((priority, value))
            self.paragraph = None

def json_articles(value):
    if isinstance(value, list):
        for item in value: yield from json_articles(item)
    elif isinstance(value, dict):
        types = value.get('@type', [])
        if isinstance(types, str): types = [types]
        if any(t in ('NewsArticle', 'Article', 'ReportageNewsArticle', 'BlogPosting') for t in types): yield value
        for key in ('@graph', 'mainEntity'):
            if key in value: yield from json_articles(value[key])

def clean_paragraphs(values):
    result = []; seen = set()
    noise = ('te puede interesar', 'lee tambien', 'leer tambien', 'suscribete', 'aviso de privacidad', 'todos los derechos reservados', 'comparte esta nota')
    for value in values:
        value = re.sub(r'\s+', ' ', text(value)).strip()
        key = normalize(value)
        if len(value.split()) < 8 or any(key.startswith(n) for n in noise) or key in seen: continue
        result.append(value); seen.add(key)
    return result

def extract_article(html):
    parser = NewsParser(); parser.feed(html); parser._finish_paragraph()
    title = normalize(' '.join(parser.page_title))
    if any(x in title for x in ('access denied', 'just a moment', 'captcha', 'pagina no encontrada', 'page not found')):
        raise ValueError('El portal mostró una página de bloqueo o error.')
    for script in parser.scripts:
        try:
            for article in json_articles(json.loads(script)):
                if text(article.get('articleBody')):
                    body = re.sub(r'<[^>]+>', ' ', text(article['articleBody']))
                    if len(body.split()) >= 30: return body, 'Texto de la noticia'
        except (ValueError, TypeError):
            continue
    paragraphs = clean_paragraphs([p for priority, p in parser.paragraphs if priority == 2])
    if not paragraphs:
        # Generic article tags are also used for unrelated recommendation cards.
        # Keep only paragraphs matching the actual page headline, not merely the portal name.
        headline = set(re.findall(r'\w+', title)) - STOP - {'puebla', 'noticias', 'azteca', 'arrow', 'up'}
        paragraphs = clean_paragraphs([p for priority, p in parser.paragraphs
            if len((set(re.findall(r'\w+', normalize(p))) - STOP) & headline) >= min(2, max(1, len(headline)))])
    if paragraphs: return '\n'.join(paragraphs), 'Texto de la noticia'
    # A description is a real portal excerpt, not text inferred from the headline.
    descriptions = clean_paragraphs(parser.descriptions)
    if descriptions: return descriptions[0], 'Descripción publicada por el portal'
    raise ValueError('No se encontró texto de la noticia. Pega el contenido o escribe el resumen.')

STOP = set('de del la el los las un una unos unas a al en y o por para con que se su sus como es fue son ha ante este esta sobre lo no'.split())
def make_excerpt(content, title='', max_words=75):
    """Elige oraciones que coinciden con el título; no añade hechos ni redacta con IA."""
    paragraphs = [re.sub(r'\s+', ' ', p).strip() for p in text(content).splitlines() if p.strip()]
    sentences = [s for p in paragraphs for s in re.split(r'(?<=[.!?])\s+(?=[A-ZÁÉÍÓÚÜÑ¿¡0-9])', p)]
    sentences = [s for s in sentences if len(s.split()) >= 5]
    if not sentences: return ''
    title_words = set(re.findall(r'\w+', normalize(title))) - STOP
    # Keep the search near the lead, avoiding unrelated stories at the end of a page.
    candidates = sentences[:14]
    def score(i):
        words = set(re.findall(r'\w+', normalize(candidates[i]))) - STOP
        return len(words & title_words) * 3 + (3 if i == 0 else 0) + (1 if re.search(r'\d', candidates[i]) else 0) - i * .18
    chosen = []; budget = 0
    for i in sorted(range(len(candidates)), key=score, reverse=True):
        words = set(re.findall(r'\w+', normalize(candidates[i]))) - STOP
        if chosen and not (words & title_words):
            continue
        count = len(candidates[i].split())
        if count <= max_words - budget:
            chosen.append(i); budget += count
        if len(chosen) == 2: break
    if chosen: return ' '.join(candidates[i] for i in sorted(chosen))
    # One long lead: preserve wording, visibly mark truncation.
    return ' '.join(candidates[0].split()[:max_words]).rstrip(' ,;:') + '…'

def check_public_url(url):
    if not valid_url(url): raise ValueError('El enlace no es una URL HTTP o HTTPS válida.')
    parsed = urlsplit(url)
    if parsed.username or parsed.password or parsed.port not in (None, 80, 443):
        raise ValueError('El enlace debe usar un portal público sin credenciales.')
    try:
        addresses = socket.getaddrinfo(parsed.hostname, parsed.port or (443 if parsed.scheme == 'https' else 80), type=socket.SOCK_STREAM)
    except socket.gaierror as exc:
        raise ValueError('No se pudo resolver el portal. Revisa la conexión o completa el resumen manualmente.') from exc
    for result in addresses:
        if not ipaddress.ip_address(result[4][0]).is_global:
            raise ValueError('El enlace no corresponde a un portal público.')

class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl): return None

def fetch_excerpt(url, title=''):
    current = url
    for _ in range(5):
        check_public_url(current)
        request = Request(current, headers={'User-Agent': 'Mozilla/5.0 (compatible; ExtractorMedios/2.0)', 'Accept': 'text/html,application/xhtml+xml'})
        try:
            with build_opener(NoRedirect).open(request, timeout=12) as response:
                if 'html' not in response.headers.get_content_type():
                    raise ValueError('El enlace no devolvió una página HTML.')
                raw = response.read(4_000_001)
                if len(raw) > 4_000_000: raise ValueError('La página excede el tamaño permitido de lectura.')
                charset = response.headers.get_content_charset() or 'utf-8'
                html = raw.decode(charset, errors='replace')
            content, source = extract_article(html)
            excerpt = make_excerpt(content, title)
            if not excerpt: raise ValueError('No se pudo obtener un extracto del contenido.')
            return {'Resumen': excerpt, 'Estado': source}
        except HTTPError as exc:
            if exc.code in (301, 302, 303, 307, 308) and exc.headers.get('Location'):
                current = urljoin(current, exc.headers['Location']); continue
            raise ValueError(f'El portal respondió HTTP {exc.code}. Completa el resumen manualmente.') from exc
    raise ValueError('El portal redirigió demasiadas veces.')
