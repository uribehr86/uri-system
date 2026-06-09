import zipfile, re, os, sys

def strip_mail_merge(path):
    try:
        if not zipfile.is_zipfile(path):
            return
        tmp = path + '.tmp'
        with zipfile.ZipFile(path, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == 'word/settings.xml':
                        xml = data.decode('utf-8')
                        xml = re.sub(r'<w:mailMerge>.*?</w:mailMerge>', '', xml, flags=re.DOTALL)
                        data = xml.encode('utf-8')
                    zout.writestr(item, data)
        os.replace(tmp, path)
        print(f'Cleaned: {os.path.basename(path)}', flush=True)
    except Exception as e:
        try:
            if os.path.exists(path + '.tmp'):
                os.remove(path + '.tmp')
        except:
            pass

dirs = [r'C:\Users\uri\Desktop', r'C:\Users\uri\Desktop\test']
for d in dirs:
    if os.path.exists(d):
        for f in os.listdir(d):
            if f.endswith('.docx') and not f.startswith('~$'):
                strip_mail_merge(os.path.join(d, f))

print('Done!', flush=True)
