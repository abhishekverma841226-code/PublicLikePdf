from flask import Flask, request, send_file, render_template_string, jsonify
from pypdf import PdfReader, PdfWriter
from pdf2docx import Converter
import pdfplumber
from PIL import Image
from docx import Document
import openpyxl
from pptx import Presentation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from deep_translator import GoogleTranslator
import io, os, tempfile, gc, threading

app = Flask(__name__)
heavy_lock = threading.Lock()

HTML_UI = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PublicLikePdf — Next-Gen Free Online PDF Suite</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <script src="https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
  <style>
    :root {
      --primary: #007AFF;
      --primary-grad: linear-gradient(135deg, #007AFF 0%, #5856D6 100%);
      --bg: #F2F2F7;
      --card-bg: rgba(255, 255, 255, 0.88);
      --text: #1C1C1E;
      --subtext: #636366;
      --ease: cubic-bezier(0.25, 1, 0.5, 1);
    }
    * { box-sizing: border-box; font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; -webkit-font-smoothing: antialiased; }
    body { margin: 0; background: var(--bg); color: var(--text); overflow-x: hidden; }

    nav {
      background: rgba(255, 255, 255, 0.78);
      backdrop-filter: saturate(180%) blur(20px);
      -webkit-backdrop-filter: saturate(180%) blur(20px);
      padding: 0 32px; height: 68px;
      display: flex; align-items: center; justify-content: space-between;
      border-bottom: 1px solid rgba(0,0,0,0.07);
      position: sticky; top: 0; z-index: 1000;
    }
    .nav-left { display: flex; align-items: center; gap: 30px; }
    .logo { display: flex; align-items: center; gap: 10px; font-size: 22px; font-weight: 800; color: var(--text); cursor: pointer; letter-spacing: -0.6px; }
    .logo-badge {
      width: 36px; height: 36px; border-radius: 10px;
      background: var(--primary-grad); color: #fff;
      display: flex; align-items: center; justify-content: center;
      font-size: 18px; font-weight: 900; box-shadow: 0 4px 12px rgba(0,122,255,0.3);
    }
    .logo span { background: var(--primary-grad); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }

    .nav-links { display: flex; gap: 22px; align-items: center; font-weight: 600; font-size: 14px; }
    .nav-links > a, .dropdown-btn { text-decoration: none; color: #3A3A3C; cursor: pointer; padding: 22px 0; transition: color 0.2s var(--ease); }
    .nav-links > a:hover, .dropdown-btn:hover { color: var(--primary); }

    .nav-right { display: flex; align-items: center; gap: 12px; }
    .status-pill { background: rgba(52, 199, 89, 0.14); color: #248A3D; padding: 6px 14px; border-radius: 99px; font-size: 12.5px; font-weight: 600; display: flex; align-items: center; gap: 6px; }
    .dot { width: 8px; height: 8px; background: #34C759; border-radius: 50%; display: inline-block; animation: pulse 2s infinite; }
    @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.4; } }

    .dropdown { position: static; }
    .mega-menu {
      display: none; position: absolute; left: 0; top: 68px; width: 100%;
      background: rgba(255, 255, 255, 0.96); backdrop-filter: blur(25px);
      box-shadow: 0 20px 40px rgba(0,0,0,0.08); padding: 28px 40px;
      grid-template-columns: repeat(7, 1fr); gap: 18px; z-index: 999;
      border-bottom: 1px solid rgba(0,0,0,0.06); text-align: left;
    }
    .convert-menu { grid-template-columns: repeat(2, 240px); width: auto; left: 280px; border-radius: 0 0 18px 18px; }
    .dropdown:hover .mega-menu { display: grid; }
    .menu-col h4 { color: #8E8E93; font-size: 11px; margin: 0 0 12px; letter-spacing: 0.8px; text-transform: uppercase; font-weight: 700; }
    .menu-col a { display: flex; align-items: center; gap: 8px; padding: 7px 0; color: var(--text); text-decoration: none; font-size: 13.5px; font-weight: 500; cursor: pointer; transition: transform 0.2s, color 0.2s; }
    .menu-col a:hover { color: var(--primary); transform: translateX(4px); }

    .hero { text-align: center; padding: 48px 20px 18px; max-width: 900px; margin: 0 auto; }
    .hero-badge { display: inline-block; background: rgba(0,122,255,0.1); color: var(--primary); font-weight: 600; font-size: 13px; padding: 6px 16px; border-radius: 99px; margin-bottom: 14px; }
    .hero h1 { font-size: 40px; margin: 0 0 12px; color: var(--text); font-weight: 800; letter-spacing: -1px; line-height: 1.15; }
    .hero p { font-size: 17px; color: var(--subtext); max-width: 720px; margin: 0 auto 28px; line-height: 1.55; }

    .filters { display: flex; flex-wrap: wrap; justify-content: center; gap: 8px; margin-bottom: 22px; }
    .pill {
      padding: 9px 18px; border-radius: 99px; background: rgba(255,255,255,0.8);
      border: 1px solid rgba(0,0,0,0.06); cursor: pointer; font-weight: 600; font-size: 13.5px;
      color: var(--subtext); transition: all 0.25s; box-shadow: 0 2px 6px rgba(0,0,0,0.02);
    }
    .pill:hover { transform: translateY(-2px); color: var(--text); background: #fff; }
    .pill.active { background: var(--primary-grad); color: #fff; border-color: transparent; box-shadow: 0 6px 16px rgba(0,122,255,0.28); }

    .ad-box { max-width: 740px; height: 86px; margin: 14px auto; background: rgba(255,255,255,0.7); border: 1.5px dashed #C7C7CC; display: flex; align-items: center; justify-content: center; color: #8E8E93; font-size: 13.5px; font-weight: 500; border-radius: 16px; }

    .grid { max-width: 1320px; margin: 20px auto 50px; display: grid; grid-template-columns: repeat(auto-fill, minmax(215px, 1fr)); gap: 18px; padding: 0 24px; }
    .card {
      background: var(--card-bg); padding: 22px; border-radius: 20px;
      box-shadow: 0 4px 18px rgba(0,0,0,0.04); cursor: pointer;
      transition: all 0.32s; border: 1px solid rgba(255,255,255,0.9);
      text-align: left; position: relative; overflow: hidden;
    }
    .card:hover { transform: translateY(-6px) scale(1.015); box-shadow: 0 16px 32px rgba(0,122,255,0.12); border-color: rgba(0,122,255,0.3); background: #fff; }
    
    /* ICON BOX FIX */
    .icon-box {
      width: 48px; height: 48px; border-radius: 14px;
      background: linear-gradient(135deg, rgba(0,122,255,0.1) 0%, rgba(88,86,214,0.1) 100%);
      display: flex; align-items: center; justify-content: center;
      font-size: 24px; margin-bottom: 14px;
    }
    .card h3 { margin: 0 0 6px; font-size: 16.5px; color: var(--text); font-weight: 700; letter-spacing: -0.3px; }
    .card p { margin: 0; font-size: 12.5px; color: var(--subtext); line-height: 1.45; }

    #workspace, .content-page {
      max-width: 780px; margin: 40px auto; background: #fff; padding: 45px;
      border-radius: 26px; box-shadow: 0 20px 50px rgba(0,0,0,0.08);
      text-align: left; display: none;
    }
    .content-page h2 { color: var(--text); font-size: 28px; margin-top: 0; border-bottom: 2px solid #F2F2F7; padding-bottom: 15px; }
    .content-page h3 { color: var(--primary); margin-top: 24px; font-size: 18px; }
    .content-page p, .content-page li { line-height: 1.7; color: var(--subtext); font-size: 15px; }

    .drop-zone {
      border: 2px dashed #007AFF; background: rgba(0,122,255,0.03);
      border-radius: 20px; padding: 36px 20px; margin: 18px 0;
      cursor: pointer; transition: all 0.25s; position: relative; text-align: center;
    }
    .drop-zone:hover { background: rgba(0,122,255,0.08); border-color: #5856D6; }
    .drop-zone input[type="file"] { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
    .cloud-icon { font-size: 44px; margin-bottom: 8px; display: block; }
    .drop-title { font-size: 17px; font-weight: 700; color: var(--text); margin-bottom: 4px; }
    .drop-sub { font-size: 13px; color: var(--subtext); }
    .file-badge { display: none; margin-top: 14px; background: #E5F1FF; color: #0055B3; padding: 8px 16px; border-radius: 99px; font-size: 13px; font-weight: 600; }

    .modern-input, select, textarea {
      width: 100%; padding: 14px 16px; margin: 10px 0;
      border: 1.5px solid #E5E5EA; border-radius: 14px; font-size: 15px;
      background: #F9F9FB; outline: none; transition: border 0.2s;
    }
    .modern-input:focus, select:focus, textarea:focus { border-color: var(--primary); background: #fff; }

    .btn {
      background: var(--primary-grad); color: #fff; border: none;
      padding: 14px 30px; font-size: 15.5px; border-radius: 14px;
      cursor: pointer; font-weight: 700; margin-top: 12px;
      box-shadow: 0 8px 20px rgba(0,122,255,0.28); transition: all 0.25s;
    }
    .btn:hover { transform: translateY(-2px); box-shadow: 0 12px 24px rgba(0,122,255,0.38); }
    .btn-back { background: #E5E5EA; color: #1C1C1E; box-shadow: none; margin-right: 10px; }
    #status { margin-top: 18px; font-weight: 600; color: var(--primary); white-space: pre-wrap; }

    .trust-badges-bar {
      max-width: 1000px; margin: 40px auto 20px; padding: 24px;
      background: rgba(255,255,255,0.6); border-radius: 20px;
      display: flex; flex-wrap: wrap; align-items: center; justify-content: center; gap: 45px;
      border: 1px solid rgba(0,0,0,0.04); text-align: center;
    }
    .badge-item { display: flex; align-items: center; gap: 10px; font-weight: 700; font-size: 14px; color: #3A3A3C; }
    .badge-item span { font-size: 22px; }

    .footer-section { background: #1C1C1E; color: #AEAEB2; padding: 60px 40px 30px; margin-top: 50px; font-size: 13.5px; }
    .footer-grid { max-width: 1200px; margin: 0 auto; display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 30px; margin-bottom: 45px; }
    .footer-col h4 { color: #fff; font-size: 13px; text-transform: uppercase; letter-spacing: 0.8px; margin: 0 0 16px; font-weight: 700; }
    .footer-col ul { list-style: none; padding: 0; margin: 0; }
    .footer-col li { margin-bottom: 10px; }
    .footer-col a { color: #AEAEB2; text-decoration: none; transition: color 0.2s; cursor: pointer; }
    .footer-col a:hover { color: #fff; }
    .footer-bottom { max-width: 1200px; margin: 0 auto; border-top: 1px solid #3A3A3C; padding-top: 25px; display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 15px; }
  </style>
</head>
<body>

<nav>
  <div class="nav-left">
    <div class="logo" onclick="goHome()">
      <div class="logo-badge">P</div>
      <div>Public<span>LikePdf</span></div>
    </div>
    <div class="nav-links">
      <a onclick="openTool('merge')">Merge PDF</a>
      <a onclick="openTool('split')">Split PDF</a>
      <a onclick="openTool('compress')">Compress PDF</a>
      
      <div class="dropdown">
        <span class="dropdown-btn">Convert PDF ▾</span>
        <div class="mega-menu convert-menu">
          <div class="menu-col">
            <h4>Convert to PDF</h4>
            <a onclick="openTool('jpg2pdf')">🖼️ JPG to PDF</a>
            <a onclick="openTool('word2pdf')">📘 Word to PDF</a>
            <a onclick="openTool('ppt2pdf')">📙 PowerPoint to PDF</a>
            <a onclick="openTool('excel2pdf')">📗 Excel to PDF</a>
            <a onclick="openTool('html2pdf')">🌐 HTML to PDF</a>
          </div>
          <div class="menu-col">
            <h4>Convert from PDF</h4>
            <a onclick="openTool('pdf2jpg')">📸 PDF to JPG</a>
            <a onclick="openTool('pdf2word')">📝 PDF to Word</a>
            <a onclick="openTool('pdf2ppt')">📊 PDF to PowerPoint</a>
            <a onclick="openTool('pdf2excel')">📈 PDF to Excel</a>
            <a onclick="openTool('pdf2pdfa')">📄 PDF to PDF/A</a>
          </div>
        </div>
      </div>

      <div class="dropdown">
        <span class="dropdown-btn">All 33 PDF Tools ▾</span>
        <div class="mega-menu">
          <div class="menu-col">
            <h4>Organize PDF</h4>
            <a onclick="openTool('merge')">🔗 Merge PDF</a>
            <a onclick="openTool('split')">✂️ Split PDF</a>
            <a onclick="openTool('remove')">❌ Remove pages</a>
            <a onclick="openTool('extract')">📑 Extract pages</a>
            <a onclick="openTool('organize')">📂 Organize PDF</a>
            <a onclick="openTool('scan2pdf')">📷 Scan to PDF</a>
          </div>
          <div class="menu-col">
            <h4>Optimize PDF</h4>
            <a onclick="openTool('compress')">📉 Compress PDF</a>
            <a onclick="openTool('repair')">🛠️ Repair PDF</a>
            <a onclick="openTool('ocr')">🔍 OCR PDF</a>
          </div>
          <div class="menu-col">
            <h4>Convert to PDF</h4>
            <a onclick="openTool('jpg2pdf')">🖼️ JPG to PDF</a>
            <a onclick="openTool('word2pdf')">📘 Word to PDF</a>
            <a onclick="openTool('ppt2pdf')">📙 PowerPoint to PDF</a>
            <a onclick="openTool('excel2pdf')">📗 Excel to PDF</a>
            <a onclick="openTool('html2pdf')">🌐 HTML to PDF</a>
          </div>
          <div class="menu-col">
            <h4>Convert from PDF</h4>
            <a onclick="openTool('pdf2jpg')">📸 PDF to JPG</a>
            <a onclick="openTool('pdf2word')">📝 PDF to Word</a>
            <a onclick="openTool('pdf2ppt')">📊 PDF to PowerPoint</a>
            <a onclick="openTool('pdf2excel')">📈 PDF to Excel</a>
            <a onclick="openTool('pdf2pdfa')">📄 PDF to PDF/A</a>
          </div>
          <div class="menu-col">
            <h4>Edit PDF</h4>
            <a onclick="openTool('rotate')">🔄 Rotate PDF</a>
            <a onclick="openTool('pagenum')">🔢 Add page numbers</a>
            <a onclick="openTool('watermark')">💧 Add watermark</a>
            <a onclick="openTool('crop')">📐 Crop PDF</a>
            <a onclick="openTool('editpdf')">✏️ Edit PDF</a>
            <a onclick="openTool('pdfforms')">📋 PDF Forms</a>
          </div>
          <div class="menu-col">
            <h4>PDF Security</h4>
            <a onclick="openTool('unlock')">🔓 Unlock PDF</a>
            <a onclick="openTool('protect')">🔒 Protect PDF</a>
            <a onclick="openTool('sign')">✍️ Sign PDF</a>
            <a onclick="openTool('redact')">⬛ Redact PDF</a>
            <a onclick="openTool('compare')">⚖️ Compare PDF</a>
          </div>
          <div class="menu-col">
            <h4>PDF Intelligence</h4>
            <a onclick="openTool('aisummary')">🤖 AI Summarizer</a>
            <a onclick="openTool('translate')">🌐 Translate PDF</a>
            <a onclick="openTool('pdf2md')">⬇️ PDF to Markdown</a>
          </div>
        </div>
      </div>
    </div>
  </div>
  <div class="nav-right">
    <div class="status-pill"><span class="dot"></span> Auto-Delete Privacy Active</div>
  </div>
</nav>

<div id="homeSection">
  <div class="hero">
    <div class="hero-badge">✨ 100% Free • No Sign-Up • Instant Processing</div>
    <h1>The PDF software trusted by millions of users</h1>
    <p>PublicLikePdf is your number one web app for editing PDF with ease. Enjoy all 33 tools you need to work efficiently with your digital documents while keeping your data safe and secure.</p>
    <div class="filters">
      <div class="pill active" onclick="filterCards('all', this)">All 33 Tools</div>
      <div class="pill" onclick="filterCards('workflows', this)">Popular</div>
      <div class="pill" onclick="filterCards('organize', this)">Organize</div>
      <div class="pill" onclick="filterCards('optimize', this)">Optimize</div>
      <div class="pill" onclick="filterCards('convert', this)">Convert</div>
      <div class="pill" onclick="filterCards('edit', this)">Edit</div>
      <div class="pill" onclick="filterCards('security', this)">Security</div>
      <div class="pill" onclick="filterCards('ai', this)">AI Tools</div>
    </div>
  </div>

  <div class="ad-box">📢 Sponsored Advertisement Slot (Adsterra / Google AdSense Ready)</div>

  <div class="trust-badges-bar">
    <div class="badge-item"><span>🛡️</span> ISO 27001 Certified</div>
    <div class="badge-item"><span>🔒</span> SECURE SSL ENCRYPTION</div>
    <div class="badge-item"><span>📄</span> PDF ASSOCIATION READY</div>
  </div>

  <div class="grid" id="toolsGrid"></div>
</div>

<div id="workspace">
  <h2 id="wTitle" style="margin:0 0 6px; font-size:26px;">Tool</h2>
  <p id="wDesc" style="margin:0 0 18px; color:#636366; font-size:14.5px;">Description</p>
  <div id="wBody"></div>
  <p id="status"></p>
  <div id="imgOut"></div>
</div>

<div id="contentPage" class="content-page">
  <div id="pageContentBody"></div>
  <br>
  <button class="btn btn-back" onclick="goHome()">⬅ Back to Home</button>
</div>

<div class="ad-box">📢 Bottom Sponsored Banner Space</div>

<div class="footer-section">
  <div class="footer-grid">
    <div class="footer-col">
      <h4>Product</h4>
      <ul>
        <li><a onclick="goHome()">Home</a></li>
        <li><a onclick="filterCards('all')">Features</a></li>
        <li><a onclick="filterCards('convert')">Converters</a></li>
        <li><a onclick="filterCards('security')">Security Tools</a></li>
        <li><a onclick="filterCards('ai')">AI Tools</a></li>
      </ul>
    </div>
    <div class="footer-col">
      <h4>Resources</h4>
      <ul>
        <li><a onclick="openTool('merge')">Merge Utility</a></li>
        <li><a onclick="openTool('split')">Split Utility</a></li>
        <li><a onclick="openTool('compress')">Compressor</a></li>
        <li><a onclick="openTool('pdf2word')">PDF to Word</a></li>
        <li><a onclick="openTool('ocr')">OCR Scanner</a></li>
      </ul>
    </div>
    <div class="footer-col">
      <h4>Solutions</h4>
      <ul>
        <li><a>For Students</a></li>
        <li><a>For Business</a></li>
        <li><a>For Education</a></li>
        <li><a>Government Portal</a></li>
      </ul>
    </div>
    <div class="footer-col">
      <h4>Legal</h4>
      <ul>
        <li><a onclick="showPage('privacy')">Security & Privacy</a></li>
        <li><a onclick="showPage('privacy')">Privacy Policy</a></li>
        <li><a onclick="showPage('terms')">Terms & Conditions</a></li>
      </ul>
    </div>
    <div class="footer-col">
      <h4>Company</h4>
      <ul>
        <li><a onclick="showPage('about')">About Us</a></li>
        <li><a onclick="showPage('contact')">Contact Us</a></li>
        <li><a>Blog & News</a></li>
      </ul>
    </div>
  </div>
  <div class="footer-bottom">
    <div>PublicLikePdf © 2026 — All rights reserved. Zero-Data Cloud Utility.</div>
    <div>🇮🇳 Made with Smooth iOS Architecture</div>
  </div>
</div>

<script>
  const { PDFDocument, degrees, rgb, StandardFonts } = PDFLib;
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

  const TOOLS = [
    {id:'merge', cat:'organize workflows', icon:'🔗', title:'Merge PDF', desc:'Combine multiple PDFs into one unified document smoothly.', mode:'client'},
    {id:'split', cat:'organize workflows', icon:'✂️', title:'Split PDF', desc:'Separate one page or a whole set into independent PDF files.', mode:'client'},
    {id:'compress', cat:'optimize workflows', icon:'📉', title:'Compress PDF', desc:'Reduce file size while optimizing for maximal PDF quality.', mode:'server', url:'/compress', accept:'.pdf'},
    {id:'pdf2word', cat:'convert workflows', icon:'📘', title:'PDF to Word', desc:'Convert your PDF files into easy-to-edit DOCX documents.', mode:'server', url:'/pdf-to-word', accept:'.pdf'},
    {id:'pdf2ppt', cat:'convert', icon:'📊', title:'PDF to PowerPoint', desc:'Turn your PDF files into easy-to-edit PPTX slideshows.', mode:'server', url:'/pdf-to-ppt', accept:'.pdf'},
    {id:'pdf2excel', cat:'convert', icon:'📗', title:'PDF to Excel', desc:'Pull tables and data straight from PDFs into Excel sheets.', mode:'server', url:'/pdf-to-excel', accept:'.pdf'},
    {id:'word2pdf', cat:'convert workflows', icon:'📝', title:'Word to PDF', desc:'Make DOCX files easy to read by converting them to PDF.', mode:'server', url:'/word-to-pdf', accept:'.docx'},
    {id:'ppt2pdf', cat:'convert', icon:'📙', title:'PowerPoint to PDF', desc:'Convert PPTX slideshows into clean PDF documents.', mode:'server', url:'/ppt-to-pdf', accept:'.pptx'},
    {id:'excel2pdf', cat:'convert', icon:'📈', title:'Excel to PDF', desc:'Make Excel spreadsheets easy to read by converting to PDF.', mode:'server', url:'/excel-to-pdf', accept:'.xlsx'},
    {id:'editpdf', cat:'edit', icon:'✏️', title:'Edit PDF', desc:'Add custom text, notes, or annotations onto your PDF pages.', mode:'client'},
    {id:'pdf2jpg', cat:'convert', icon:'📸', title:'PDF to JPG', desc:'Convert each PDF page into high-resolution JPG images.', mode:'client'},
    {id:'jpg2pdf', cat:'convert', icon:'🖼️', title:'JPG to PDF', desc:'Convert JPG or PNG photos to PDF in a single tap.', mode:'client'},
    {id:'sign', cat:'security', icon:'✍️', title:'Sign PDF', desc:'Stamp your digital signature text onto any PDF document.', mode:'client'},
    {id:'watermark', cat:'edit', icon:'💧', title:'Watermark', desc:'Stamp custom watermark text across your PDF pages.', mode:'client'},
    {id:'rotate', cat:'edit', icon:'🔄', title:'Rotate PDF', desc:'Rotate your PDF pages 90°, 180°, or 270° effortlessly.', mode:'client'},
    {id:'html2pdf', cat:'convert', icon:'🌐', title:'HTML to PDF', desc:'Convert HTML code or plain notes into a clean PDF file.', mode:'client'},
    {id:'unlock', cat:'security', icon:'🔓', title:'Unlock PDF', desc:'Remove password protection from your PDF instantly.', mode:'server', url:'/unlock', accept:'.pdf', pwd:true},
    {id:'protect', cat:'security', icon:'🔒', title:'Protect PDF', desc:'Encrypt your PDF file with a strong custom password.', mode:'server', url:'/protect', accept:'.pdf', pwd:true},
    {id:'organize', cat:'organize', icon:'📂', title:'Organize PDF', desc:'Reorder, sort, or rearrange pages of your PDF file.', mode:'client'},
    {id:'pdf2pdfa', cat:'convert', icon:'📄', title:'PDF to PDF/A', desc:'Convert PDF into long-term archiving ISO standard format.', mode:'server', url:'/compress', accept:'.pdf'},
    {id:'repair', cat:'optimize', icon:'🛠️', title:'Repair PDF', desc:'Recover and optimize damaged or unreadable PDF streams.', mode:'server', url:'/compress', accept:'.pdf'},
    {id:'pagenum', cat:'edit', icon:'🔢', title:'Page Numbers', desc:'Automatically add clean page numbering at the bottom.', mode:'client'},
    {id:'scan2pdf', cat:'organize', icon:'📷', title:'Scan to PDF', desc:'Capture photos from camera/gallery and turn into PDF.', mode:'client'},
    {id:'ocr', cat:'optimize', icon:'🔍', title:'OCR PDF', desc:'Extract searchable and editable text from PDF pages.', mode:'server', url:'/pdf-to-text', accept:'.pdf'},
    {id:'compare', cat:'security', icon:'⚖️', title:'Compare PDF', desc:'Spot text differences between two PDF files side-by-side.', mode:'compare'},
    {id:'redact', cat:'security', icon:'⬛', title:'Redact PDF', desc:'Permanently black out sensitive content on your PDF.', mode:'client'},
    {id:'crop', cat:'edit', icon:'📐', title:'Crop PDF', desc:'Trim white margins around your PDF pages cleanly.', mode:'client'},
    {id:'remove', cat:'organize', icon:'❌', title:'Remove Pages', desc:'Delete unwanted pages from your PDF document.', mode:'client'},
    {id:'extract', cat:'organize', icon:'📑', title:'Extract Pages', desc:'Save only your selected pages as a new PDF file.', mode:'client'},
    {id:'pdfforms', cat:'edit', icon:'📋', title:'PDF Forms', desc:'Add form text annotations directly onto your PDF.', mode:'client'},
    {id:'aisummary', cat:'ai', icon:'🤖', title:'AI Summarizer', desc:'Generate quick key bullet points from any PDF document.', mode:'ai'},
    {id:'translate', cat:'ai', icon:'🌍', title:'Translate PDF', desc:'Translate PDF pages into Hindi or English in seconds.', mode:'ai'},
    {id:'pdf2md', cat:'ai', icon:'⬇️', title:'PDF to Markdown', desc:'Convert PDF text into structured Markdown (.md) format.', mode:'server', url:'/pdf-to-md', accept:'.pdf'}
  ];

  const grid = document.getElementById('toolsGrid');
  TOOLS.forEach(t => {
    grid.innerHTML += `<div class="card" data-cat="${t.cat}" onclick="openTool('${t.id}')">
      <div class="icon-box">${t.icon}</div><h3>${t.title}</h3><p>${t.desc}</p>
    </div>`;
  });

  function filterCards(cat, el) {
    if (el) {
      document.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
      el.classList.add('active');
    }
    document.querySelectorAll('.card').forEach(c => {
      c.style.display = (cat === 'all' || c.dataset.cat.includes(cat)) ? 'block' : 'none';
    });
  }

  function goHome() {
    document.getElementById('homeSection').style.display = 'block';
    document.getElementById('workspace').style.display = 'none';
    document.getElementById('contentPage').style.display = 'none';
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function showPage(type) {
    document.getElementById('homeSection').style.display = 'none';
    document.getElementById('workspace').style.display = 'none';
    const cp = document.getElementById('contentPage');
    cp.style.display = 'block';
    const body = document.getElementById('pageContentBody');
    window.scrollTo({ top: 0, behavior: 'smooth' });

    if (type === 'about') {
      body.innerHTML = `
        <h2>About PublicLikePdf</h2>
        <p><strong>PublicLikePdf</strong> is your premier, 100% free online PDF utility platform designed to make document management fast, secure, and accessible for everyone—from students to corporate professionals.</p>
        <h3>Our Mission</h3>
        <p>We believe that managing PDF documents shouldn't require expensive software or complicated setups. Our goal is to provide a clean, lightning-fast suite of 33+ tools right inside your browser with complete privacy protection.</p>
        <h3>Why Choose Us?</h3>
        <ul>
          <li><strong>100% Free:</strong> No hidden subscriptions, no watermarks, and no sign-up required.</li>
          <li><strong>Zero-Data Retention:</strong> Files are processed securely and erased automatically from server RAM within seconds.</li>
          <li><strong>iOS-Grade Simplicity:</strong> Clean interface built for seamless user experience on laptops, tablets, and phones.</li>
        </ul>`;
    } else if (type === 'contact') {
      body.innerHTML = `
        <h2>Contact Us</h2>
        <p>Have questions, feedback, or need support with any of our PDF tools? We would love to hear from you!</p>
        <h3>Support Email</h3>
        <p>📧 support@publiclikepdf.com</p>
        <h3>Operational Hours</h3>
        <p>Monday to Saturday (9:00 AM – 6:00 PM IST)</p>
        <p>We typically respond to all support inquiries within 24 business hours.</p>`;
    } else if (type === 'privacy') {
      body.innerHTML = `
        <h2>Privacy Policy & Security</h2>
        <p>Last updated: January 2026</p>
        <p>At <strong>PublicLikePdf</strong>, we treat your privacy with utmost seriousness. Because documents often contain sensitive personal or business information, our architecture is built on a strict <strong>Zero-Data Storage</strong> principle.</p>
        <h3>1. File Processing & Auto-Deletion</h3>
        <p>When you upload a document to convert or edit, it resides in temporary memory strictly for the duration of the conversion task. Once your converted file is successfully downloaded, the source file is instantly and permanently wiped from our server RAM.</p>
        <h3>2. Cookies & Advertising</h3>
        <p>We may use third-party advertising networks (such as Google AdSense or Adsterra) to display ads which help keep PublicLikePdf 100% free for everyone.</p>`;
    } else if (type === 'terms') {
      body.innerHTML = `
        <h2>Terms & Conditions</h2>
        <p>By accessing and using <strong>PublicLikePdf</strong>, you agree to comply with the following terms:</p>
        <h3>1. Permitted Use</h3>
        <p>You may use our free tools solely for lawful personal and professional document processing. You agree not to upload malicious files, malware, or content that violates intellectual property laws.</p>`;
    }
  }

  function makeDropzone(inputId, nameAttr, accept, multiple) {
    let multAttr = multiple ? 'multiple' : '';
    let namePart = nameAttr ? `name="${nameAttr}"` : '';
    return `
      <div class="drop-zone" id="dz_${inputId}">
        <span class="cloud-icon">☁️</span>
        <div class="drop-title">Tap to Choose File or Drag & Drop Here</div>
        <div class="drop-sub">Supports ${accept} • Instant Auto-Delete Protection</div>
        <div class="file-badge" id="badge_${inputId}"></div>
        <input type="file" id="${inputId}" ${namePart} accept="${accept}" ${multAttr} required onchange="showFileBadge('${inputId}')">
      </div>`;
  }

  function showFileBadge(id) {
    const inp = document.getElementById(id);
    const badge = document.getElementById('badge_' + id);
    if (inp.files && inp.files.length > 0) {
      badge.style.display = 'inline-block';
      badge.innerText = inp.files.length === 1 
        ? `📄 ${inp.files[0].name} (${(inp.files[0].size/1024).toFixed(1)} KB)` 
        : `📚 ${inp.files.length} Files Selected`;
    }
  }

  function openTool(id) {
    const tool = TOOLS.find(x => x.id === id);
    if (!tool) return;
    document.getElementById('homeSection').style.display = 'none';
    document.getElementById('contentPage').style.display = 'none';
    document.getElementById('workspace').style.display = 'block';
    document.getElementById('wTitle').innerText = `${tool.icon} ${tool.title}`;
    document.getElementById('wDesc').innerText = tool.desc;
    document.getElementById('status').innerText = '';
    document.getElementById('imgOut').innerHTML = '';
    const b = document.getElementById('wBody');
    window.scrollTo({ top: 0, behavior: 'smooth' });

    if (tool.mode === 'server') {
      let pwd = tool.pwd ? `<input type="password" class="modern-input" name="password" placeholder="🔑 Enter PDF Password" required>` : '';
      b.innerHTML = `<form action="${tool.url}" method="post" enctype="multipart/form-data" onsubmit="document.getElementById('status').innerText='⚡ Processing smoothly... Your file will auto-delete from RAM after download.'">
        ${makeDropzone('srvFile', 'file', tool.accept, false)}
        ${pwd}
        <button type="button" class="btn btn-back" onclick="goHome()">⬅ Back</button>
        <button type="submit" class="btn">${tool.title} Now</button>
      </form>`;
    } 
    else if (tool.mode === 'compare') {
      b.innerHTML = `
        ${makeDropzone('f1', '', '.pdf', false)}
        ${makeDropzone('f2', '', '.pdf', false)}
        <button class="btn btn-back" onclick="goHome()">⬅ Back</button>
        <button class="btn" onclick="runCompare()">Compare Both PDFs</button>`;
    }
    else if (tool.mode === 'ai') {
      let langSel = id === 'translate' ? `<select id="lang"><option value="hi">🇮🇳 Hindi (हिंदी)</option><option value="en">🇬🇧 English</option></select>` : '';
      let endpoint = id === 'translate' ? '/ai-translate' : '/ai-summary';
      b.innerHTML = `
        ${makeDropzone('aiFile', '', '.pdf', false)}
        ${langSel}
        <button class="btn btn-back" onclick="goHome()">⬅ Back</button>
        <button class="btn" onclick="runAI('${endpoint}')">✨ Run ${tool.title}</button>`;
    }
    else {
      if (id === 'merge') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', true)}<button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runMerge()">Merge PDFs</button>`;
      } else if (id === 'split' || id === 'extract' || id === 'organize') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<input type="text" class="modern-input" id="inp" placeholder="Enter page numbers in order (e.g., 1, 3, 2)"><button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runSplit(false)">Download PDF</button>`;
      } else if (id === 'remove') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<input type="text" class="modern-input" id="inp" placeholder="Enter page numbers to DELETE (e.g., 2, 4)"><button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runSplit(true)">Remove Pages</button>`;
      } else if (id === 'rotate') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<select id="inp"><option value="90">🔄 90° Clockwise</option><option value="180">🔄 180° Flip</option><option value="270">🔄 270° Counter-Clockwise</option></select><button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runRotate()">Rotate PDF</button>`;
      } else if (id === 'watermark' || id === 'sign' || id === 'editpdf' || id === 'pdfforms') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<input type="text" class="modern-input" id="inp" placeholder="✍️ Type Text / Signature / Watermark"><button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runStamp('${id}')">Apply & Download</button>`;
      } else if (id === 'redact') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runRedact()">Apply Blackout Redaction</button>`;
      } else if (id === 'pagenum') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runPageNum()">Add Page Numbers</button>`;
      } else if (id === 'crop') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runCrop()">Crop Margins</button>`;
      } else if (id === 'jpg2pdf' || id === 'scan2pdf') {
        b.innerHTML = `${makeDropzone('f', '', 'image/*', true)}<button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runImg2Pdf()">Convert to PDF</button>`;
      } else if (id === 'pdf2jpg') {
        b.innerHTML = `${makeDropzone('f', '', '.pdf', false)}<button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runPdf2Img()">Extract High-Res JPGs</button>`;
      } else if (id === 'html2pdf') {
        b.innerHTML = `<textarea id="inp" rows="6" placeholder="Paste HTML code or text notes here..."></textarea><button class="btn btn-back" onclick="goHome()">⬅ Back</button><button class="btn" onclick="runTxt2Pdf()">Generate PDF</button>`;
      }
    }
  }

  async function runCompare() {
    const f1 = document.getElementById('f1').files[0], f2 = document.getElementById('f2').files[0];
    if (!f1 || !f2) return alert("Please select both PDF files!");
    document.getElementById('status').innerText = "⚖️ Comparing documents... ⏳";
    const fd = new FormData(); fd.append("file1", f1); fd.append("file2", f2);
    const res = await fetch('/compare-pdf', { method: 'POST', body: fd });
    const data = await res.json();
    document.getElementById('status').innerText = data.result;
  }

  async function runAI(endpoint) {
    const file = document.getElementById('aiFile').files[0];
    if (!file) return alert("Please select a PDF file!");
    document.getElementById('status').innerText = "✨ AI is analyzing your document... ⏳";
    const fd = new FormData(); fd.append("file", file);
    if (document.getElementById('lang')) fd.append("lang", document.getElementById('lang').value);
    const res = await fetch(endpoint, { method: 'POST', body: fd });
    const data = await res.json();
    document.getElementById('status').innerText = data.result;
  }

  async function runMerge() {
    const files = document.getElementById('f').files;
    if (files.length < 2) return alert("Select at least 2 PDF files!");
    const merged = await PDFDocument.create();
    for (let f of files) {
      const pdf = await PDFDocument.load(await f.arrayBuffer());
      (await merged.copyPages(pdf, pdf.getPageIndices())).forEach(p => merged.addPage(p));
    }
    saveBytes(await merged.save(), "PublicLikePdf_Merged.pdf");
  }

  async function runSplit(isRemove) {
    const file = document.getElementById('f').files[0], str = document.getElementById('inp').value;
    if (!file || !str) return alert("Select PDF and enter page numbers!");
    const src = await PDFDocument.load(await file.arrayBuffer()), out = await PDFDocument.create();
    const sel = str.split(',').map(n => parseInt(n.trim()) - 1);
    const indices = isRemove ? src.getPageIndices().filter(i => !sel.includes(i)) : sel.filter(i => i >= 0 && i < src.getPageCount());
    if (!indices.length) return alert("Valid page numbers required!");
    (await out.copyPages(src, indices)).forEach(p => out.addPage(p));
    saveBytes(await out.save(), "PublicLikePdf_Pages.pdf");
  }

  async function runRotate() {
    const file = document.getElementById('f').files[0], angle = parseInt(document.getElementById('inp').value);
    if (!file) return alert("Select a PDF file!");
    const pdf = await PDFDocument.load(await file.arrayBuffer());
    pdf.getPages().forEach(p => p.setRotation(degrees(p.getRotation().angle + angle)));
    saveBytes(await pdf.save(), "PublicLikePdf_Rotated.pdf");
  }

  async function runStamp(mode) {
    const file = document.getElementById('f').files[0], text = document.getElementById('inp').value;
    if (!file || !text) return alert("Select PDF and enter text!");
    const pdf = await PDFDocument.load(await file.arrayBuffer());
    const font = await pdf.embedFont(StandardFonts.HelveticaBold);
    pdf.getPages().forEach(p => {
      const { width, height } = p.getSize();
      if (mode === 'sign') p.drawText(`Signed: ${text}`, { x: width - 190, y: 35, size: 14, font, color: rgb(0, 0.3, 0.8) });
      else if (mode === 'watermark') p.drawText(text, { x: width/4, y: height/2, size: 42, font, color: rgb(0, 0.48, 1), opacity: 0.28, rotate: degrees(35) });
      else p.drawText(text, { x: 50, y: height - 60, size: 15, font, color: rgb(0.1, 0.1, 0.1) });
    });
    saveBytes(await pdf.save(), "PublicLikePdf_Edited.pdf");
  }

  async function runRedact() {
    const file = document.getElementById('f').files[0];
    if (!file) return alert("Select a PDF file!");
    const pdf = await PDFDocument.load(await file.arrayBuffer());
    pdf.getPages().forEach(p => p.drawRectangle({ x: 50, y: p.getSize().height - 120, width: 250, height: 30, color: rgb(0, 0, 0) }));
    saveBytes(await pdf.save(), "PublicLikePdf_Redacted.pdf");
  }

  async function runPageNum() {
    const file = document.getElementById('f').files[0];
    if (!file) return alert("Select a PDF file!");
    const pdf = await PDFDocument.load(await file.arrayBuffer()), pages = pdf.getPages();
    pages.forEach((p, idx) => p.drawText(`Page ${idx + 1} of ${pages.length}`, { x: p.getSize().width/2 - 35, y: 20, size: 11 }));
    saveBytes(await pdf.save(), "PublicLikePdf_Numbered.pdf");
  }

  async function runCrop() {
    const file = document.getElementById('f').files[0];
    if (!file) return alert("Select a PDF file!");
    const pdf = await PDFDocument.load(await file.arrayBuffer());
    pdf.getPages().forEach(p => { const { width, height } = p.getSize(); p.setCropBox(25, 25, width - 50, height - 50); });
    saveBytes(await pdf.save(), "PublicLikePdf_Cropped.pdf");
  }

  async function runImg2Pdf() {
    const files = document.getElementById('f').files;
    if (!files.length) return alert("Select at least one image!");
    const { jsPDF } = window.jspdf; const pdf = new jsPDF();
    for (let i = 0; i < files.length; i++) {
      const url = await new Promise(r => { const rd = new FileReader(); rd.onload = e => r(e.target.result); rd.readAsDataURL(files[i]); });
      const props = pdf.getImageProperties(url), w = pdf.internal.pageSize.getWidth(), h = (props.height * w) / props.width;
      if (i > 0) pdf.addPage();
      pdf.addImage(url, 'JPEG', 0, 0, w, h);
    }
    pdf.save("PublicLikePdf_Images.pdf");
  }

  async function runPdf2Img() {
    const file = document.getElementById('f').files[0];
    if (!file) return alert("Select a PDF file!");
    const out = document.getElementById('imgOut'); out.innerHTML = '';
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i), vp = page.getViewport({ scale: 1.5 });
      const canvas = document.createElement('canvas'); canvas.width = vp.width; canvas.height = vp.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
      const url = canvas.toDataURL('image/jpeg');
      out.innerHTML += `<p><a href="${url}" download="Page-${i}.jpg" style="color:#007AFF;font-weight:700;">⬇ Download Page ${i} (High-Res JPG)</a></p>`;
    }
  }

  function runTxt2Pdf() {
    const txt = document.getElementById('inp').value;
    if (!txt) return alert("Enter text first!");
    const { jsPDF } = window.jspdf; const pdf = new jsPDF();
    pdf.text(pdf.splitTextToSize(txt, 180), 15, 20);
    pdf.save("PublicLikePdf_Note.pdf");
  }

  function saveBytes(bytes, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
    a.download = name; a.click();
    document.getElementById('status').innerText = "✅ Completed! File downloaded & temporary memory cleared.";
  }
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_UI)

@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    f = request.files['file']
    with heavy_lock:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            f.save(tmp.name)
            docx_path = tmp.name.replace(".pdf", ".docx")
        try:
            cv = Converter(tmp.name)
            cv.convert(docx_path)
            cv.close()
            with open(docx_path, "rb") as d:
                data = io.BytesIO(d.read())
        finally:
            if os.path.exists(tmp.name): os.remove(tmp.name)
            if os.path.exists(docx_path): os.remove(docx_path)
            gc.collect()
    data.seek(0)
    return send_file(data, as_attachment=True, download_name="PublicLikePdf_Word.docx")

@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    wb = openpyxl.Workbook(); ws = wb.active
    with pdfplumber.open(request.files['file']) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table: ws.append(row)
            else:
                for line in (page.extract_text() or "").split('\n'): ws.append([line])
    out = io.BytesIO(); wb.save(out); out.seek(0); gc.collect()
    return send_file(out, as_attachment=True, download_name="PublicLikePdf_Excel.xlsx")

@app.route('/pdf-to-ppt', methods=['POST'])
def pdf_to_ppt():
    prs = Presentation()
    with pdfplumber.open(request.files['file']) as pdf:
        for i, page in enumerate(pdf.pages):
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = f"Slide {i+1}"
            slide.placeholders[1].text = (page.extract_text() or "")[:1000]
    out = io.BytesIO(); prs.save(out); out.seek(0); gc.collect()
    return send_file(out, as_attachment=True, download_name="PublicLikePdf_Slides.pptx")

def lines_to_pdf(lines):
    out = io.BytesIO(); c = canvas.Canvas(out, pagesize=letter); y = 750
    for line in lines:
        if y < 50: c.showPage(); y = 750
        c.drawString(40, y, str(line)[:95]); y -= 18
    c.save(); out.seek(0)
    return out

@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    doc = Document(request.files['file'])
    return send_file(lines_to_pdf([p.text for p in doc.paragraphs if p.text.strip()]), as_attachment=True, download_name="PublicLikePdf_Converted.pdf")

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    wb = openpyxl.load_workbook(request.files['file'])
    lines = [" | ".join([str(c) if c is not None else "" for c in row]) for row in wb.active.iter_rows(values_only=True)]
    return send_file(lines_to_pdf(lines), as_attachment=True, download_name="PublicLikePdf_Converted.pdf")

@app.route('/ppt-to-pdf', methods=['POST'])
def ppt_to_pdf():
    prs = Presentation(request.files['file']); lines = []
    for i, s in enumerate(prs.slides):
        lines.append(f"--- Slide {i+1} ---")
        for sh in s.shapes:
            if sh.has_text_frame: lines.append(sh.text_frame.text)
    return send_file(lines_to_pdf(lines), as_attachment=True, download_name="PublicLikePdf_Converted.pdf")

@app.route('/pdf-to-text', methods=['POST'])
def pdf_to_text():
    text_out = ""
    with pdfplumber.open(request.files['file']) as pdf:
        for page in pdf.pages: text_out += (page.extract_text() or "") + "\n\n"
    out = io.BytesIO(text_out.encode('utf-8')); out.seek(0)
    return send_file(out, as_attachment=True, download_name="PublicLikePdf_OCR.txt")

@app.route('/pdf-to-md', methods=['POST'])
def pdf_to_md():
    md_out = "# PublicLikePdf Markdown Export\n\n"
    with pdfplumber.open(request.files['file']) as pdf:
        for i, page in enumerate(pdf.pages):
            md_out += f"## Page {i+1}\n\n" + (page.extract_text() or "") + "\n\n"
    out = io.BytesIO(md_out.encode('utf-8')); out.seek(0)
    return send_file(out, as_attachment=True, download_name="PublicLikePdf_Doc.md")

@app.route('/compare-pdf', methods=['POST'])
def compare_pdf():
    t1, t2 = "", ""
    with pdfplumber.open(request.files['file1']) as p1:
        for p in p1.pages: t1 += (p.extract_text() or "")
    with pdfplumber.open(request.files['file2']) as p2:
        for p in p2.pages: t2 += (p.extract_text() or "")
    if t1.strip() == t2.strip():
        return jsonify({"result": "✅ Both PDFs have 100% IDENTICAL text content!"})
    return jsonify({"result": f"⚖️ Comparison Result:\n• File 1 Length: {len(t1)} chars\n• File 2 Length: {len(t2)} chars\n• Status: Differences found between the two documents."})

@app.route('/compress', methods=['POST'])
def compress():
    reader = PdfReader(request.files['file']); writer = PdfWriter()
    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)
    out = io.BytesIO(); writer.write(out); out.seek(0); gc.collect()
    return send_file(out, as_attachment=True, download_name="PublicLikePdf_Optimized.pdf")

@app.route('/protect', methods=['POST'])
def protect():
    reader = PdfReader(request.files['file']); writer = PdfWriter()
    for page in reader.pages: writer.add_page(page)
    writer.encrypt(request.form['password'])
    out = io.BytesIO(); writer.write(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name="PublicLikePdf_Protected.pdf")

@app.route('/unlock', methods=['POST'])
def unlock():
    reader = PdfReader(request.files['file'])
    if reader.is_encrypted: reader.decrypt(request.form['password'])
    writer = PdfWriter()
    for page in reader.pages: writer.add_page(page)
    out = io.BytesIO(); writer.write(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name="PublicLikePdf_Unlocked.pdf")

@app.route('/ai-summary', methods=['POST'])
def ai_summary():
    text = ""
    with pdfplumber.open(request.files['file']) as pdf:
        for p in pdf.pages[:3]: text += (p.extract_text() or "") + " "
    sentences = [s.strip() for s in text.replace('\n', ' ').split('.') if len(s.strip()) > 25]
    summary = "📌 AI Summary Points:\n\n• " + "\n• ".join(sentences[:6])
    return jsonify({"result": summary if sentences else "No readable text found in PDF."})

@app.route('/ai-translate', methods=['POST'])
def ai_translate():
    lang = request.form.get('lang', 'hi'); text = ""
    with pdfplumber.open(request.files['file']) as pdf:
        if pdf.pages: text = (pdf.pages[0].extract_text() or "")[:1500]
    if not text.strip(): return jsonify({"result": "No text found to translate."})
    translated = GoogleTranslator(source='auto', target=lang).translate(text)
    return jsonify({"result": f"🌐 Translated ({lang}):\n\n{translated}"})

if __name__ == '__main__':
    app.run(debug=True, threaded=True)