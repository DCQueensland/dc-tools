#!/usr/bin/env python3
"""
Dale Carnegie QLD - Program Material Finder
Scans entire Dropbox for training materials and lets you search
by project brief / outcomes / pain points to find relevant modules,
visuals, participant manuals, and trainer guides.
"""

import os
import json
import re
import http.server
import socketserver
import urllib.parse
from collections import defaultdict
from datetime import datetime

PORT = 8767
DROPBOX_ROOT = os.path.expanduser("~/Dropbox")
EXTENSIONS = {'.pptx', '.ppt', '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.png', '.jpg', '.jpeg', '.mp4', '.mov'}
EXCLUDE_DIRS = {'.dropbox.cache', 'Icon\r', '.dropbox'}
EXCLUDE_PATTERNS = {'~$', '.DS_Store'}

# File type classification
FILE_CATEGORIES = {
    'visuals': ['visuals', 'slides', 'powerpoint slides', 'presentation'],
    'participant_manual': ['pm', 'participant manual', 'participant', 'manual', 'handout', 'workbook'],
    'trainer_manual': ['trainer manual', 'trainer', 'facilitator guide', 'facilitator', 'trainer guide', 'tm'],
    'breakthrough': ['breakthrough', 'breakthrough sheet'],
    'video': ['.mp4', '.mov'],
    'template': ['template'],
    'proposal': ['proposal'],
    'flyer': ['flyer'],
    'certificate': ['certificate'],
    'packing_list': ['packing list', 'kit checklist'],
}

# Topic keyword mappings for smarter search
TOPIC_KEYWORDS = {
    'leadership': ['leadership', 'leader', 'leading', 'ltm', 'uylp', 'ltr', 'delegation', 'coaching'],
    'communication': ['communication', 'communicating', 'presenting', 'presentation', 'hip', 'public speaking', 'assertive'],
    'sales': ['sales', 'selling', 'wrs', 'prospecting', 'objections', 'closing', 'relationship selling', 'gain access'],
    'confidence': ['confidence', 'self confidence', 'self-confidence', 'fear', 'comfort zone', 'dcc'],
    'teamwork': ['team', 'teams', 'hpt', 'high performance', 'collaboration', 'collaborative', 'team building'],
    'customer_service': ['customer service', 'wccs', 'complaint', 'telephone', 'service', 'customer value'],
    'stress': ['stress', 'pressure', 'resilience', 'wellbeing', 'burnout', 'demands'],
    'conflict': ['conflict', 'disagree', 'difficult conversations', 'frank conversations', 'negotiation'],
    'change': ['change', 'adapt', 'flexibility', 'transition', 'influence change', 'leading change'],
    'engagement': ['engagement', 'motivation', 'motivate', 'recognition', 'inspire', 'morale'],
    'innovation': ['innovation', 'creative', 'problem solving', 'decision making', 'process improvement'],
    'relationships': ['relationships', 'networking', 'rapport', 'trust', 'names', 'memory', 'connect'],
    'coaching': ['coaching', 'mentoring', 'feedback', 'performance', 'appraisal', 'prd'],
    'meetings': ['meetings', 'effective meetings', 'facilitation'],
    'time_management': ['time management', 'priorities', 'planning', 'delegation', 'multiple demands'],
    'personality': ['personality', 'disc', 'personality styles', 'communication styles', 'behavioural'],
    'habits': ['habits', 'wgyh', 'what got you here', 'goldsmith', 'hold you back', 'paradox'],
    'influence': ['influence', 'persuasion', 'buy-in', 'cooperation', 'winning people'],
    'school': ['school', 'principal', 'teacher', 'student', 'youth', 'gen next', 'shs', 'education'],
}


def scan_dropbox():
    """Scan entire Dropbox and build file index."""
    files = []
    for root, dirs, filenames in os.walk(DROPBOX_ROOT):
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
        for fname in filenames:
            if any(fname.startswith(p) for p in EXCLUDE_PATTERNS):
                continue
            ext = os.path.splitext(fname)[1].lower()
            if ext not in EXTENSIONS:
                continue
            full_path = os.path.join(root, fname)
            rel_path = os.path.relpath(full_path, DROPBOX_ROOT)
            try:
                mtime = os.path.getmtime(full_path)
                size = os.path.getsize(full_path)
            except OSError:
                mtime = 0
                size = 0

            fname_lower = fname.lower()
            path_lower = rel_path.lower()
            file_type = classify_file(fname_lower, path_lower, ext)

            search_text = rel_path.replace('/', ' ').replace('\\', ' ').replace('-', ' ').replace('_', ' ')
            search_text = re.sub(r'\.[a-z]+$', '', search_text).lower()

            files.append({
                'path': rel_path,
                'full_path': full_path,
                'name': fname,
                'ext': ext,
                'type': file_type,
                'search_text': search_text,
                'mtime': mtime,
                'size': size,
                'folder': os.path.dirname(rel_path),
            })
    return files


def classify_file(fname_lower, path_lower, ext):
    """Classify a file into a category based on its name and path."""
    if ext in ('.mp4', '.mov'):
        return 'video'
    if ext in ('.png', '.jpg', '.jpeg'):
        return 'image'

    for category, keywords in FILE_CATEGORIES.items():
        for kw in keywords:
            if kw.startswith('.'):
                continue
            if kw in fname_lower:
                return category

    if 'visual' in path_lower or 'slides' in path_lower:
        return 'visuals'
    if ext == '.pptx' or ext == '.ppt':
        return 'presentation'
    if ext == '.pdf':
        return 'document_pdf'
    if ext in ('.docx', '.doc'):
        return 'document_word'
    if ext in ('.xlsx', '.xls'):
        return 'spreadsheet'
    return 'other'


def search_files(file_index, query):
    """Search files using query keywords with topic expansion."""
    query_lower = query.lower()

    query_words = set(re.findall(r'[a-z]+(?:\s+[a-z]+)?', query_lower))

    expanded_keywords = set()
    for word in query_words:
        expanded_keywords.add(word)
        for topic, synonyms in TOPIC_KEYWORDS.items():
            for syn in synonyms:
                if word in syn or syn in word or word == topic:
                    expanded_keywords.update(synonyms)

    all_search_terms = set()
    for kw in expanded_keywords:
        all_search_terms.add(kw)
        for part in kw.split():
            if len(part) > 2:
                all_search_terms.add(part)

    scored = []
    for f in file_index:
        score = 0
        matched_terms = set()
        st = f['search_text']

        for word in query_words:
            if word in st and len(word) > 2:
                score += 10
                matched_terms.add(word)

        query_phrases = re.findall(r'[a-z]+(?:\s+[a-z]+)+', query_lower)
        for phrase in query_phrases:
            if phrase in st:
                score += 20
                matched_terms.add(phrase)

        for term in all_search_terms:
            if term in st and term not in query_words:
                score += 3
                matched_terms.add(term)

        if 'all programs & workshops breakdowns' in st:
            score += 5
        if 'class management' in st:
            score += 4
        if 'product file' in st:
            score += 3

        if f['type'] in ('visuals', 'participant_manual', 'trainer_manual'):
            score += 3
        if f['type'] == 'breakthrough':
            score += 2

        if f['mtime'] > datetime(2025, 1, 1).timestamp():
            score += 1
        if f['mtime'] > datetime(2026, 1, 1).timestamp():
            score += 1

        if f['name'].startswith('~$'):
            score -= 100

        if score >= 8:
            scored.append({
                **f,
                'score': score,
                'matched_terms': list(matched_terms),
            })

    scored.sort(key=lambda x: (-x['score'], -x['mtime']))
    return scored


def group_results(results, max_results=300):
    """Group results by folder/module for cleaner display."""
    groups = defaultdict(list)
    seen = set()

    for r in results[:max_results]:
        folder = r['folder']
        if r['path'] not in seen:
            seen.add(r['path'])
            groups[folder].append(r)

    sorted_groups = sorted(
        groups.items(),
        key=lambda x: (
            max(f['score'] for f in x[1]),
            max(f['mtime'] for f in x[1]),
        ),
        reverse=True
    )
    return sorted_groups


# Build the index on startup
print("Scanning Dropbox for training materials...")
APP_STATE = {'index': scan_dropbox()}
print(f"Indexed {len(APP_STATE['index'])} files across Dropbox.")


class ProgramFinderHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == '/' or parsed.path == '':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'program-finder.html')
            with open(html_path, 'rb') as f:
                self.wfile.write(f.read())

        elif parsed.path == '/api/search':
            params = urllib.parse.parse_qs(parsed.query)
            query = params.get('q', [''])[0]

            if not query.strip():
                self.send_json({'groups': [], 'total': 0})
                return

            results = search_files(APP_STATE['index'], query)
            groups = group_results(results)

            json_groups = []
            for folder, files in groups:
                json_groups.append({
                    'folder': folder,
                    'files': [{
                        'path': f['path'],
                        'name': f['name'],
                        'ext': f['ext'],
                        'type': f['type'],
                        'score': f['score'],
                        'matched_terms': f['matched_terms'],
                        'size': f['size'],
                        'mtime': f['mtime'],
                    } for f in files]
                })

            self.send_json({
                'groups': json_groups,
                'total': len(results),
                'query': query,
            })

        elif parsed.path == '/api/open':
            params = urllib.parse.parse_qs(parsed.query)
            filepath = params.get('path', [''])[0]
            if filepath:
                full_path = os.path.join(DROPBOX_ROOT, filepath)
                if os.path.exists(full_path):
                    os.system(f'open "{full_path}"')
                    self.send_json({'ok': True})
                else:
                    self.send_json({'ok': False, 'error': 'File not found'})
            else:
                self.send_json({'ok': False, 'error': 'No path'})

        elif parsed.path == '/api/open-folder':
            params = urllib.parse.parse_qs(parsed.query)
            filepath = params.get('path', [''])[0]
            if filepath:
                full_path = os.path.join(DROPBOX_ROOT, filepath)
                folder = os.path.dirname(full_path)
                if os.path.exists(folder):
                    os.system(f'open "{folder}"')
                    self.send_json({'ok': True})
                else:
                    self.send_json({'ok': False, 'error': 'Folder not found'})
            else:
                self.send_json({'ok': False, 'error': 'No path'})

        elif parsed.path == '/api/rescan':
            APP_STATE['index'] = scan_dropbox()
            self.send_json({'ok': True, 'count': len(APP_STATE['index'])})

        elif parsed.path == '/api/stats':
            ext_counts = defaultdict(int)
            type_counts = defaultdict(int)
            for f in APP_STATE['index']:
                ext_counts[f['ext']] += 1
                type_counts[f['type']] += 1
            self.send_json({
                'total_files': len(APP_STATE['index']),
                'by_extension': dict(ext_counts),
                'by_type': dict(type_counts),
            })

        else:
            self.send_response(404)
            self.end_headers()

    def send_json(self, data):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())


if __name__ == '__main__':
    with socketserver.TCPServer(("", PORT), ProgramFinderHandler) as httpd:
        httpd.allow_reuse_address = True
        print(f"Program Material Finder running at http://localhost:{PORT}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down.")
