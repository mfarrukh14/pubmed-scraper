#!/usr/bin/env python3
"""
Article extractor with Tkinter GUI (append mode) and custom Gene column formatting.

Features:
 - Scrapes PubMed/journal pages for metadata and abstract heuristics.
 - Extracts first author and places it in 'Author(s)'.
 - Appends each extraction to a chosen Excel file (preserves existing rows).
 - Formats the 'Gene' column as a numbered list:
     1:PPARG(rs1801282 (Pro12Ala))
     2:HNF4A(rs745975)
     3:GLIS3(rs6415788, rs806052)
     ...
   Each numbered item is placed on its own line in the Excel cell.
"""

import re, os, html, datetime
import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import socket
import threading
import time

# ---------- Configuration ----------
GENOTYPE_KEYWORDS = ['MassARRAY','massarray','MassARRAY genotyping','whole exome sequencing','WES','PCR','TaqMan','microarray','sequencing','Sanger','genotyping']
STUDY_DESIGN_KEYWORDS = {
    'Case control': ['case-control','case control','case-control study'],
    'Cohort': ['cohort','prospective cohort','retrospective cohort'],
    'Cross-sectional': ['cross-sectional'],
    'RCT': ['randomized','randomised','randomized controlled trial'],
    'Meta-analysis': ['meta-analysis','meta analysis','systematic review']
}
REGION_KEYWORDS = ['Pakistan','Pashtun','KPK','Khyber','Sindh','Punjab','Balochistan','India','China','USA','United States']
README_STATUS_URL = "https://github.com/mfarrukh14/pubmed-scraper/blob/main/README.md"

# ---------- Permission and Connection Checks ----------
def check_internet_connection():
    """Check if internet connection is available"""
    try:
        # Try to connect to Google's DNS server
        socket.create_connection(("8.8.8.8", 53), timeout=3)
        return True
    except OSError:
        return False

def check_application_status():
    """Check if the application is enabled by reading the GitHub README status"""
    try:
        if not check_internet_connection():
            return False, "internet"
        
        headers = {'User-Agent': 'Mozilla/5.0 (ArticleExtractor/1.0)'}
        response = requests.get(README_STATUS_URL, headers=headers, timeout=10)
        response.raise_for_status()
        
        content = response.text.strip()
        # Look for Status: ON or Status: OFF (case insensitive)
        if re.search(r'Status:\s*ON', content, re.IGNORECASE):
            return True, "enabled"
        elif re.search(r'Status:\s*OFF', content, re.IGNORECASE):
            return False, "disabled"
        else:
            # If status format is not found, assume disabled for security
            return False, "disabled"
    except Exception:
        # If README is unreachable, deny access
        return False, "unreachable"

def show_permission_dialog(parent, status_type):
    """Show appropriate permission/error dialog based on status type"""
    if status_type == "internet":
        messagebox.showerror(
            "No Internet Connection", 
            "Please connect to internet to use application",
            parent=parent
        )
    elif status_type in ["disabled", "unreachable"]:
        messagebox.showerror(
            "Permission Required", 
            "Please obtain permissions from Farrukh to use this application.",
            parent=parent
        )
    return False

# ---------- Helpers ----------
def fetch_html(url, timeout=15):
    headers = {'User-Agent':'Mozilla/5.0 (ArticleExtractor/1.0)'}
    r = requests.get(url, headers=headers, timeout=timeout)
    r.raise_for_status()
    return r.text

def soup_from_html(html_text):
    return BeautifulSoup(html_text, 'lxml')

def get_meta(soup, key):
    tag = soup.find('meta', attrs={'name': key})
    if tag and tag.get('content'):
        return tag['content'].strip()
    tag = soup.find('meta', attrs={'property': key})
    if tag and tag.get('content'):
        return tag['content'].strip()
    return None

def extract_basic_metadata(soup):
    # Title / DOI / PMID / Year (unchanged)
    title = get_meta(soup, 'citation_title') or get_meta(soup, 'og:title') or (soup.title.string.strip() if soup.title else '')
    journal = get_meta(soup, 'citation_journal_title') or get_meta(soup, 'citation_journal') or get_meta(soup, 'og:site_name') or ''
    doi = get_meta(soup, 'citation_doi') or get_meta(soup, 'DC.identifier') or ''
    pmid = get_meta(soup, 'citation_pmid') or ''
    date = get_meta(soup, 'citation_publication_date') or get_meta(soup, 'citation_date') or ''
    year = ''
    if date:
        m = re.search(r'(\d{4})', date)
        if m: year = m.group(1)
    if not year:
        t = soup.get_text(' ')
        m = re.search(r'©\s*(\d{4})', t)
        if m: year = m.group(1)

    # --- Robust author handling: prefer meta tags always ---
    authors_meta = [tag.get('content').strip() for tag in soup.find_all('meta', attrs={'name': 'citation_author'}) if tag.get('content')]
    first_author = authors_meta[0] if authors_meta else ''
    authors_str = ', '.join(authors_meta) if authors_meta else ''

    # Fallback: conservative PubMed-style lookup only if meta absent
    if not first_author:
        # PubMed often exposes authors as links; try common containers but avoid sweeping page text
        candidate = soup.select_one('.authors-list, .authors, #authors, .author-list, .full-authors')
        if candidate:
            # pick the first reasonable anchor/span/li that looks like a person name
            for tag in candidate.find_all(['a','span','li'], limit=30):
                txt = tag.get_text(' ', strip=True)
                if txt and re.match(r'^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+$', txt):
                    first_author = txt
                    break
        # final fallback: first anchor that looks like "Name Surname" but only among first 120 anchors
        if not first_author:
            for a in soup.find_all('a', limit=120):
                txt = a.get_text(' ', strip=True)
                if txt and re.match(r'^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+$', txt):
                    first_author = txt
                    break
        if not authors_str and first_author:
            authors_str = first_author

    return {
        'title': html.unescape(title) if title else '',
        'first_author': first_author,
        'authors_list': authors_str or first_author,
        'journal': html.unescape(journal) if journal else '',
        'doi': doi or '',
        'pmid': pmid or '',
        'year': year or ''
    }


def get_abstract_text(soup):
    candidate = get_meta(soup, 'citation_abstract') or get_meta(soup, 'og:description') or get_meta(soup, 'description')
    if candidate: return candidate
    # find a container with class/id containing 'abstract'
    abs_tag = soup.find(lambda t: t.name in ['div','section','p'] and (('abstract' in ' '.join(t.get('class', [])).lower()) or ('abstract' in (t.get('id') or '').lower())))
    if abs_tag: return abs_tag.get_text(' ', strip=True)
    # fallback: largest paragraph
    paras = soup.find_all('p')
    if paras:
        pbest = max(paras, key=lambda p: len(p.get_text()))
        return pbest.get_text(' ', strip=True)
    return ''

# ---------- Heuristic extractors ----------
def find_study_design(text):
    if not text: return ''
    t = text.lower()
    for label, keys in STUDY_DESIGN_KEYWORDS.items():
        for k in keys:
            if k in t:
                return label
    return ''

def find_region(text):
    if not text: return ''
    for r in REGION_KEYWORDS:
        if re.search(r'\b' + re.escape(r) + r'\b', text, re.I):
            return r
    return ''

def find_sample_size(text):
    if not text: return ''
    t = text.replace(',', '')
    m = re.search(r'\b[nN]\s*=\s*(\d{2,6})', t)
    if m: return int(m.group(1))
    m = re.search(r'sample size (?:of )?(\d{2,6})', t)
    if m: return int(m.group(1))
    m = re.search(r'(\d{2,6})\s+T2DM cases and controls each', t)
    if m: return int(m.group(1)) * 2
    m = re.search(r'(\d{2,6})\s+(?:cases|subjects)[^\d]{0,30}?(\d{2,6})\s+controls', t)
    if m:
        try: return int(m.group(1)) + int(m.group(2))
        except: pass
    nums = [int(x) for x in re.findall(r'\b(\d{2,6})\b', t)]
    if nums:
        big = max(nums)
        if big >= 30: return big
    return ''

def find_mean_age(text):
    if not text: return ''
    m = re.search(r'mean age[^0-9]{0,6}([0-9]{1,3}(?:\.[0-9]+)?)\s*(?:±|\+/-)\s*([0-9]{1,3}(?:\.[0-9]+)?)', text, re.I)
    if m: return f"{m.group(1)} ± {m.group(2)}"
    m = re.search(r'mean age (?:was|of)?\s*([0-9]{1,3}(?:\.[0-9]+)?)', text, re.I)
    if m: return m.group(1)
    m = re.search(r'case[:\s]*([0-9]{1,3}(?:\.[0-9]+)?)\s*(?:±|\+/-)\s*([0-9]{1,3}(?:\.[0-9]+)?)\s*.*control[:\s]*([0-9]{1,3}(?:\.[0-9]+)?)\s*(?:±|\+/-)\s*([0-9]{1,3}(?:\.[0-9]+)?)', text, re.I)
    if m:
        return f"case:{m.group(1)}±{m.group(2)} control:{m.group(3)}±{m.group(4)}"
    return ''

def extract_rsids_and_genes_with_annotations(text):
    """
    Returns:
      rsids: list of rsIDs in order of appearance
      gene_map: { rsid: {'gene': gene_name_or_None, 'annot': annotation_string_or_empty} }
      genes_ordered: list of gene names in order of first appearance
    """
    if not text:
        return [], {}, []

    # Find rsIDs preserving order
    rsids = list(dict.fromkeys(re.findall(r'\b(rs\d{3,})\b', text, flags=re.I)))

    gene_map = {}
    gene_first_pos = {}  # store first appearance index for ordering
    for idx, rsi in enumerate(rsids):
        gene_map[rsi] = {'gene': None, 'annot': ''}
        # Pattern 1: GENE( rsID ... ) where inside parentheses may contain extra annotation
        # e.g. "PPARG (rs1801282 (Pro12Ala))" -> capture gene 'PPARG' and inner 'rs1801282 (Pro12Ala)'
        pat1 = re.compile(r'([A-Za-z0-9\-\_]+)\s*\(\s*(' + re.escape(rsi) + r'(?:[^\)]*)\))', re.I)
        m1 = pat1.search(text)
        if m1:
            gene = m1.group(1).strip()
            inner = m1.group(2).strip()  # includes rs and inner annotation if present
            # inner may be like 'rs1801282 (Pro12Ala)' — keep as-is
            gene_map[rsi]['gene'] = gene
            gene_map[rsi]['annot'] = inner[len(rsi):].strip() if inner.startswith(rsi) else inner
            # note gene first position by index of match
            gene_first_pos.setdefault(gene, idx)
            continue
        # Pattern 2: "rs1801282/PPARG" or "rs1801282 / PPARG"
        m2 = re.search(re.escape(rsi) + r'\s*\/\s*([A-Za-z0-9\-\_]+)', text)
        if m2:
            gene = m2.group(1).strip()
            gene_map[rsi]['gene'] = gene
            gene_map[rsi]['annot'] = ''
            gene_first_pos.setdefault(gene, idx)
            continue
        # Pattern 3: uppercase gene immediately before rs (up to ~30 chars)
        m3 = re.search(r'([A-Z0-9]{2,15})\s{0,6}[^A-Za-z0-9]{0,6}' + re.escape(rsi), text)
        if m3:
            gene = m3.group(1).strip()
            gene_map[rsi]['gene'] = gene
            gene_map[rsi]['annot'] = ''
            gene_first_pos.setdefault(gene, idx)
            continue
        # else gene remains None

    # Build ordered gene list by first appearance index; include genes with at least one rs
    genes_seen = sorted(gene_first_pos.items(), key=lambda kv: kv[1])
    genes_ordered = [g for g, _ in genes_seen]
    return rsids, gene_map, genes_ordered

def find_genotyping_method(text):
    if not text: return ''
    for key in GENOTYPE_KEYWORDS:
        if key.lower() in text.lower():
            return key
    return ''

def extract_allele_freqs(text, rsids):
    freqs = {r: None for r in rsids}
    if not text: return freqs
    for rsi in rsids:
        m = re.search(re.escape(rsi) + r'[^0-9]{0,8}([01]?\.\d{1,4})', text)
        if m:
            freqs[rsi] = m.group(1)
            continue
        m = re.search(re.escape(rsi) + r'[^0-9]{0,8}([0-9]{1,3}\.\d{1,2})\s*%', text)
        if m:
            try:
                val = float(m.group(1))/100.0
                freqs[rsi] = f"{val:.3f}"
            except:
                freqs[rsi] = m.group(1) + '%'
    return freqs

def extract_or_pvals(text):
    if not text: return []
    entries = []
    pattern = re.compile(r'(?:(rs\d{3,})[^\(\),;]{0,40})?(?:OR|odds ratio)\s*[=:]?\s*([0-9]+\.[0-9]+)\b[^\)]{0,80}?(?:P\s*[=<>]\s*([0-9\.>]+))?', re.I)
    for m in pattern.finditer(text):
        entries.append((m.group(1), m.group(2), m.group(3)))
    pattern2 = re.compile(r'(rs\d{3,})[^\)]{0,40}\(.*?OR\s*[=:]?\s*([0-9]+\.[0-9]+).*?P\s*[=<>]\s*([0-9\.>]+)', re.I)
    for m in pattern2.finditer(text):
        entries.append((m.group(1), m.group(2), m.group(3)))
    unique = {}
    for rsi, orv, p in entries:
        key = rsi or f"OR_{orv}_P_{p}"
        if key not in unique:
            unique[key] = {'rsid': rsi, 'OR': orv, 'P': p}
    return list(unique.values())

def infer_effect_direction(or_str, p_str):
    try:
        orv = float(or_str)
    except:
        return ''
    p_sig = None
    if p_str:
        if '>' in p_str:
            try: p_sig = float(p_str.replace('>','')) <= 0.05
            except: p_sig = None
        else:
            try: p_sig = float(p_str) <= 0.05
            except: p_sig = None
    eff = 'Risk ↑' if orv > 1 else ('Risk ↓' if orv < 1 else 'No change')
    if p_sig is False:
        return 'No significant effect (ns)'
    return eff

# ---------- Main parse ----------
def parse_article(url):
    html_text = fetch_html(url)
    soup = soup_from_html(html_text)
    meta = extract_basic_metadata(soup)
    abstract = get_abstract_text(soup) or ''
    page_text = soup.get_text(' ')
    big = ' '.join([meta.get('title',''), meta.get('authors_list',''), meta.get('journal',''), meta.get('doi',''), abstract, page_text])

    study_design = find_study_design(big)
    region = find_region(big)
    sample_size = find_sample_size(big)
    mean_age = find_mean_age(big)
    rsids, gene_map, genes_ordered = extract_rsids_and_genes_with_annotations(big)
    genotyping_method = find_genotyping_method(big)
    allele_freqs = extract_allele_freqs(big, rsids)
    or_pvals = extract_or_pvals(big)

    reported_assoc_lines = []
    effect_dir_lines = []
    pvalue_lines = []
    for item in or_pvals:
        rsi = item.get('rsid') or ''
        orv = item.get('OR') or ''
        p = item.get('P') or ''
        label = rsi or f"OR:{orv}"
        reported_assoc_lines.append(f"{label} (OR={orv}) P={p}")
        eff = infer_effect_direction(orv, p) if orv else ''
        effect_dir_lines.append(f"{label} → {eff}" if eff else f"{label} → Unknown")
        pvalue_lines.append(f"{label} → {p}")

    if not or_pvals and rsids:
        for rsi in rsids:
            m = re.search(re.escape(rsi) + r'[^.]{0,80}P\s*[=<>]\s*([0-9\.>]+)', big, re.I)
            if m:
                pvalue_lines.append(f"{rsi} → {m.group(1)}")
                reported_assoc_lines.append(f"{rsi} → P={m.group(1)}")

    allele_freq_cell = '\n'.join([f"{gene_map.get(r,{}).get('gene','')} ({r}) → {allele_freqs.get(r)}" for r in rsids if allele_freqs.get(r)]) if rsids else ''
    snp_only_list = '\n'.join(rsids) if rsids else ''

    # ---------- Build the numbered Gene string ----------
    numbered_entries = []
    # If genes_ordered is empty, build grouping by gene_map contents or fallback per-rs
    if genes_ordered:
        idx = 1
        for gene in genes_ordered:
            # collect rs for this gene in original rsids order
            rs_for_gene = []
            for rsi in rsids:
                entry = gene_map.get(rsi, {})
                if entry.get('gene') == gene:
                    # build rs display preserving annotation if present
                    annot = entry.get('annot','').strip()
                    if annot:
                        # annot likely starts with rest of the inner parentheses after rs, e.g. ' (Pro12Ala)'
                        # We will attach it directly after rs to match example: rs1801282 (Pro12Ala)
                        # Ensure annot begins with space or '('
                        if not annot.startswith('(') and not annot.startswith(' '):
                            annot = ' ' + annot
                        rs_for_gene.append(f"{rsi}{annot}")
                    else:
                        rs_for_gene.append(rsi)
            # If no rs collected (possible), skip
            if not rs_for_gene:
                continue
            rs_list_str = ', '.join(rs_for_gene)
            numbered_entries.append(f"{idx}:{gene}({rs_list_str})")
            idx += 1
    else:
        # fallback: list each rs as its own "1:rsxxxx"
        for i, rsi in enumerate(rsids, start=1):
            numbered_entries.append(f"{i}:{rsi}")
    gene_cell = '\n'.join(numbered_entries)

    quality = ''
    if re.search(r'SIFT', big, re.I) or re.search(r'PolyPhen', big, re.I):
        quality = '1:SIFT score 2:PolyPhen score'
    comments = 'Case-control study; sample size estimation and replication required.' if 'case control' in study_design.lower() else 'Parsed heuristically from abstract/meta.'

    author_display = meta.get('first_author') or meta.get('authors_list') or ''

    row = {
        'Author(s)': author_display,
        'Title': meta.get('title',''),
        'Year': meta.get('year',''),
        'Journal': meta.get('journal',''),
        'DOI/PMID': (f"PMID: {meta.get('pmid')}\nDOI: {meta.get('doi')}").strip() if (meta.get('pmid') or meta.get('doi')) else '',
        'Study Design': study_design,
        'Region': region,
        'Sample Size (Cases)': sample_size,
        'Mean Age': mean_age,
        'Gene': gene_cell,
        'SNP/Variant': snp_only_list,
        'Genotyping Method': genotyping_method,
        'Allele Frequency (Cases)': allele_freq_cell,
        'Reported Association': '\n'.join(reported_assoc_lines),
        'Effect Direction': '\n'.join(effect_dir_lines),
        'p-value': '\n'.join(pvalue_lines),
        'Quality Score (NOS)': quality,
        'Comments/Remarks': comments
    }
    return row

# ---------- GUI & append logic ----------
DEFAULT_COLS = [
    'Author(s)','Title','Year','Journal','DOI/PMID','Study Design','Region',
    'Sample Size (Cases)','Mean Age','Gene','SNP/Variant','Genotyping Method',
    'Allele Frequency (Cases)','Reported Association','Effect Direction','p-value',
    'Quality Score (NOS)','Comments/Remarks'
]

class SplashScreen:
    def __init__(self, root):
        self.root = root
        self.splash = tk.Toplevel()
        self.splash.title("")
        self.splash.geometry("600x400")
        self.splash.configure(bg='#2c3e50')
        self.splash.resizable(False, False)
        
        # Center the splash screen
        self.splash.update_idletasks()
        width = self.splash.winfo_width()
        height = self.splash.winfo_height()
        x = (self.splash.winfo_screenwidth() // 2) - (width // 2)
        y = (self.splash.winfo_screenheight() // 2) - (height // 2)
        self.splash.geometry(f'{width}x{height}+{x}+{y}')
        
        # Remove window decorations
        self.splash.overrideredirect(True)
        
        # Create main frame
        main_frame = tk.Frame(self.splash, bg='#2c3e50', pady=50)
        main_frame.pack(expand=True, fill='both')
        
        # Add image placeholder (you can replace this with actual image)
        try:
            # Try to load an image if it exists
            if os.path.exists("ali_image.PNG") or os.path.exists("ali_image.PNG"):
                img_path = "ali_image.PNG" if os.path.exists("ali_image.PNG") else "ali_image.PNG"
                img = Image.open(img_path)
                img = img.resize((200, 200), Image.Resampling.LANCZOS)
                self.photo = ImageTk.PhotoImage(img)
                img_label = tk.Label(main_frame, image=self.photo, bg='#2c3e50')
                img_label.pack(pady=20)
            else:
                # Placeholder if no image found
                placeholder = tk.Label(main_frame, text="🎖️", font=('Arial', 60), 
                                     bg='#2c3e50', fg='#f39c12')
                placeholder.pack(pady=20)
        except Exception:
            # Fallback placeholder
            placeholder = tk.Label(main_frame, text="🎖️", font=('Arial', 60), 
                                 bg='#2c3e50', fg='#f39c12')
            placeholder.pack(pady=20)
        
        # Memorial message in italics
        memorial_text = tk.Label(main_frame, 
                                text="In loving memory of\nMaj. General Ali Nigga", 
                                font=('Times New Roman', 24, 'italic'),
                                bg='#2c3e50', 
                                fg='#ecf0f1',
                                justify='center')
        memorial_text.pack(pady=30)
        
        # Loading text
        loading_text = tk.Label(main_frame, 
                              text="Loading Application...", 
                              font=('Arial', 12),
                              bg='#2c3e50', 
                              fg='#bdc3c7')
        loading_text.pack(pady=20)
        
        # Show splash screen
        self.splash.lift()
        self.splash.focus_force()
        
        # Close splash screen after 5 seconds
        self.root.after(5000, self.close_splash)
    
    def close_splash(self):
        self.splash.destroy()

class App:
    def __init__(self, root):
        self.root = root
        root.title("Article -> Excel Extractor (Gene numbering)")
        
        # Check permissions first
        self.app_enabled = False
        self.check_permissions_on_startup()
        
        frm = ttk.Frame(root, padding=12)
        frm.grid(sticky='nsew')
        ttk.Label(frm, text="Article URL:").grid(column=0, row=0, sticky='w')
        self.url_var = tk.StringVar()
        self.entry = ttk.Entry(frm, width=95, textvariable=self.url_var)
        self.entry.grid(column=0, row=1, columnspan=4, sticky='we', pady=6)
        self.choose_btn = ttk.Button(frm, text="Choose Excel File", command=self.choose_file)
        self.choose_btn.grid(column=0, row=2, sticky='w')
        self.file_label = ttk.Label(frm, text="No file selected", foreground='blue')
        self.file_label.grid(column=1, row=2, columnspan=3, sticky='w')
        self.extract_btn = ttk.Button(frm, text="Extract & Append", command=self.on_extract)
        self.extract_btn.grid(column=0, row=3, pady=8, sticky='w')
        self.status = ttk.Label(frm, text="", foreground='green')
        self.status.grid(column=0, row=4, columnspan=4, sticky='w')
        self.excel_path = None
        
        # Disable interface if not permitted
        if not self.app_enabled:
            self.disable_interface()

    def check_permissions_on_startup(self):
        """Check permissions when the application starts"""
        self.status_checking = ttk.Label(self.root, text="Checking permissions...", foreground='orange')
        self.status_checking.pack(pady=20)
        self.root.update_idletasks()
        
        is_enabled, status_type = check_application_status()
        self.status_checking.destroy()
        
        if not is_enabled:
            show_permission_dialog(self.root, status_type)
            self.app_enabled = False
        else:
            self.app_enabled = True

    def disable_interface(self):
        """Disable all interactive elements if permissions not granted"""
        # This method will be called after interface creation if needed
        pass

    def choose_file(self):
        if not self.app_enabled:
            show_permission_dialog(self.root, "disabled")
            return
            
        f = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files','*.xlsx')], title='Choose or create Excel file to append to')
        if f:
            self.excel_path = f
            self.file_label.config(text=self.excel_path)

    def on_extract(self):
        if not self.app_enabled:
            show_permission_dialog(self.root, "disabled")
            return
            
        # Double-check permissions before extraction
        is_enabled, status_type = check_application_status()
        if not is_enabled:
            show_permission_dialog(self.root, status_type)
            return
        
        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("Error", "Paste article URL first.")
            return
        if not self.excel_path:
            messagebox.showerror("Error", "Choose or create an Excel file first.")
            return
        self.status.config(text="Fetching and parsing...")
        self.root.update_idletasks()
        try:
            row = parse_article(url)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch/parse URL:\n{e}")
            self.status.config(text=f"Failed: {e}")
            return
        df_new = pd.DataFrame([row], columns=DEFAULT_COLS)
        try:
            if os.path.exists(self.excel_path):
                try:
                    df_exist = pd.read_excel(self.excel_path, engine='openpyxl')
                except Exception:
                    df_exist = pd.DataFrame(columns=DEFAULT_COLS)
                for c in DEFAULT_COLS:
                    if c not in df_exist.columns:
                        df_exist[c] = ''
                df_combined = pd.concat([df_exist, df_new], ignore_index=True, sort=False)[DEFAULT_COLS]
            else:
                df_combined = df_new
            df_combined.to_excel(self.excel_path, index=False, engine='openpyxl')
            self.status.config(text=f"Saved/appended to: {self.excel_path}")
            messagebox.showinfo("Done", f"Saved/appended to:\n{self.excel_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")
            self.status.config(text=f"Save failed: {e}")

def main():
    root = tk.Tk()
    root.withdraw()  # Hide main window initially
    
    # Show splash screen
    splash = SplashScreen(root)
    
    # Wait for splash screen to finish (5 seconds)
    def show_main_app():
        root.deiconify()  # Show main window
        app = App(root)
    
    root.after(5000, show_main_app)
    root.mainloop()

if __name__ == '__main__':
    main()
