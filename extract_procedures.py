import zipfile
import xml.etree.ElementTree as ET
import os
import json

procedure_dir = r'c:\Users\CriZor\Documents\IFC\INFO\PROCEDURE'
temp_dir = r'c:\laragon\www\portfolio\temp_docx'
os.makedirs(temp_dir, exist_ok=True)

files = [
    'PROCEDURE Ajout d\'une machine client dans un domaine.docx',
    'PROCEDURE D\'installation Zabbix.docx',
    'PROCEDURE Profil Windows.docx',
    'PROCEDURE Routeur VLAN.docx',
    'PROCEDURE SRV-W11 (AD1).docx'
]

procedures_data = []

def extract_docx_text(docx_path):
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            xml_content = zip_ref.read('word/document.xml')
            root = ET.fromstring(xml_content)
            
            # Namespace
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            # Extract all text elements
            text_elements = []
            for paragraph in root.findall('.//w:p', ns):
                para_text = []
                for text_elem in paragraph.findall('.//w:t', ns):
                    if text_elem.text:
                        para_text.append(text_elem.text)
                if para_text:
                    text_elements.append(''.join(para_text))
            
            return '\n'.join(text_elements)
    except Exception as e:
        return f'Erreur lors de la lecture: {str(e)}'

for file in files:
    file_path = os.path.join(procedure_dir, file)
    if os.path.exists(file_path):
        text = extract_docx_text(file_path)
        # Nettoyer le nom pour le titre
        title = file.replace('.docx', '').replace('PROCEDURE ', '').strip()
        
        procedures_data.append({
            'title': title,
            'content': text,
            'category': 'Importée'
        })
        print(f'✓ {file}')
    else:
        print(f'✗ Fichier non trouvé: {file}')

# Sauvegarder en JSON
output_json = os.path.join(temp_dir, 'procedures.json')
with open(output_json, 'w', encoding='utf-8') as f:
    json.dump(procedures_data, f, ensure_ascii=False, indent=2)

print(f'\n✓ Données sauvegardées dans {output_json}')
