from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
import os
import subprocess
import platform
from datetime import datetime
from docx import Document
from docx.shared import Inches
import zipfile
import tempfile
import shutil

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this to a secure secret key

# Configuration
UPLOAD_FOLDER = 'documents'
ALLOWED_EXTENSIONS = {'docx', 'doc'}

# Create documents folder if it doesn't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_documents():
    """Get list of all Word documents in the upload folder"""
    documents = []
    for filename in os.listdir(UPLOAD_FOLDER):
        if allowed_file(filename):
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            stat = os.stat(filepath)
            documents.append({
                'name': filename,
                'size': round(stat.st_size / 1024, 2),  # Size in KB
                'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                'path': filepath
            })
    return sorted(documents, key=lambda x: x['modified'], reverse=True)

def open_document_in_word(filepath):
    """Open document in Microsoft Word or LibreOffice"""
    try:
        if platform.system() == "Windows":
            os.startfile(filepath)
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", filepath])
        else:  # Linux
            # Try different methods for Linux
            try:
                # Try LibreOffice first
                subprocess.call(["libreoffice", "--writer", filepath])
                return True
            except FileNotFoundError:
                try:
                    # Try OnlyOffice
                    subprocess.call(["onlyoffice-desktopeditors", filepath])
                    return True
                except FileNotFoundError:
                    try:
                        # Try WPS Office
                        subprocess.call(["wps", filepath])
                        return True
                    except FileNotFoundError:
                        try:
                            # Fallback to xdg-open
                            subprocess.call(["xdg-open", filepath])
                            return True
                        except:
                            return False
        return True
    except Exception as e:
        print(f"Error opening document: {e}")
        return False

@app.route('/')
def index():
    """Main page showing all documents"""
    documents = get_documents()
    return render_template('index.html', documents=documents)

@app.route('/create', methods=['GET', 'POST'])
def create_document():
    """Create a new Word document"""
    if request.method == 'POST':
        filename = request.form['filename']
        title = request.form['title']
        content = request.form['content']
        
        if not filename.endswith('.docx'):
            filename += '.docx'
        
        # Create new document
        doc = Document()
        
        # Add title
        if title:
            doc.add_heading(title, 0)
        
        # Add content
        if content:
            doc.add_paragraph(content)
        
        # Save document
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        doc.save(filepath)
        
        flash(f'Document "{filename}" created successfully!', 'success')
        return redirect(url_for('index'))
    
    return render_template('create.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload_document():
    """Upload existing Word document"""
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file selected', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash('No file selected', 'error')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = file.filename
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            
            # Handle duplicate filenames
            counter = 1
            base_name, ext = os.path.splitext(filename)
            while os.path.exists(filepath):
                filename = f"{base_name}_{counter}{ext}"
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                counter += 1
            
            file.save(filepath)
            flash(f'Document "{filename}" uploaded successfully!', 'success')
            return redirect(url_for('index'))
        else:
            flash('Invalid file type. Please upload .docx or .doc files only.', 'error')
    
    return render_template('upload.html')

@app.route('/edit/<filename>')
def edit_document(filename):
    """Open document in Microsoft Word for editing"""
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    
    if not os.path.exists(filepath):
        flash('Document not found', 'error')
        return redirect(url_for('index'))
    
    # Check if we're on a server without desktop environment
    if platform.system() == "Linux" and not os.environ.get('DISPLAY'):
        flash('Document editing is not available on this server. Please download the document to edit it locally.', 'error')
        return redirect(url_for('index'))
    
    if open_document_in_word(filepath):
        flash(f'Opening "{filename}" in document editor...', 'info')
    else:
        flash('Could not open document. Please download it and open locally, or install LibreOffice.', 'error')
    
    return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_document(filename):
    """Download document"""
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    
    if not os.path.exists(filepath):
        flash('Document not found', 'error')
        return redirect(url_for('index'))
    
    return send_file(filepath, as_attachment=True)

@app.route('/delete/<filename>')
def delete_document(filename):
    """Delete document"""
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    
    if not os.path.exists(filepath):
        flash('Document not found', 'error')
        return redirect(url_for('index'))
    
    try:
        os.remove(filepath)
        flash(f'Document "{filename}" deleted successfully!', 'success')
    except Exception as e:
        flash(f'Error deleting document: {str(e)}', 'error')
    
    return redirect(url_for('index'))

@app.route('/preview/<filename>')
def preview_document(filename):
    """Preview document content"""
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    
    if not os.path.exists(filepath):
        flash('Document not found', 'error')
        return redirect(url_for('index'))
    
    try:
        doc = Document(filepath)
        content = []
        
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                content.append(paragraph.text)
        
        return render_template('preview.html', filename=filename, content=content)
    except Exception as e:
        flash(f'Error reading document: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/webedit/<filename>', methods=['GET', 'POST'])
def web_edit_document(filename):
    """Web-based document editing"""
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    
    if not os.path.exists(filepath):
        flash('Document not found', 'error')
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        # Save the edited content
        content = request.form['content']
        title = request.form.get('title', '')
        
        # Create new document with updated content
        doc = Document()
        
        if title:
            doc.add_heading(title, 0)
        
        # Split content by lines and add as paragraphs
        for line in content.split('\n'):
            if line.strip():
                doc.add_paragraph(line)
        
        doc.save(filepath)
        flash(f'Document "{filename}" updated successfully!', 'success')
        return redirect(url_for('index'))
    
    # Read current content for editing
    try:
        doc = Document(filepath)
        content = []
        title = ""
        
        for paragraph in doc.paragraphs:
            if paragraph.style.name.startswith('Heading') and not title:
                title = paragraph.text
            elif paragraph.text.strip():
                content.append(paragraph.text)
        
        content_text = '\n'.join(content)
        return render_template('webedit.html', filename=filename, content=content_text, title=title)
    except Exception as e:
        flash(f'Error reading document: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/api/documents')
def api_documents():
    """API endpoint to get documents list"""
    return jsonify(get_documents())
    app.run(debug=True, host='0.0.0.0', port=5000)
