import zipfile, sys, os, shutil, random, string

def remove_sheet_protection(file_path):
    backup_file = file_path.replace('.xlsx', '_backup.xlsx')
    
    random_id = ''.join(
        random.choices(
            string.ascii_uppercase + string.digits, k=7
        )
    )
    
    shutil.copyfile(file_path, backup_file)
    
    zip_path = file_path.replace('.xlsx', '.zip')
    os.rename(file_path, zip_path)

    extracted_dir = 'extracted_files'
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extracted_dir)

    sheets_dir = os.path.join(extracted_dir, 'xl', 'worksheets')

    for sheet in os.listdir(sheets_dir):
        sheet_path = os.path.join(sheets_dir, sheet)
        with open(sheet_path, 'r', encoding='utf-8') as file:
            data = file.read()

        if '<sheetProtection' in data:
            data = data.replace('<sheetProtection', '<!--<sheetProtection')
            data = data.replace('/>', '/>-->')

        with open(sheet_path, 'w', encoding='utf-8') as file:
            file.write(data)

    new_zip_name = 'desbloqueado_temp.zip'
    
    with zipfile.ZipFile(new_zip_name, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, _, files in os.walk(extracted_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, extracted_dir)
                
                zip_ref.write(file_path, arcname)

    new_file_name = f'{backup_file.replace('backup', random_id)}'
    os.rename(new_zip_name, new_file_name)

    shutil.rmtree(extracted_dir)
    os.remove(zip_path)
    
    print(f"Proteção removida. O novo arquivo desbloqueado foi salvo como: {new_file_name}")
    print(f"Um backup do arquivo original foi salvo como: {backup_file}")

remove_sheet_protection(sys.argv[1])