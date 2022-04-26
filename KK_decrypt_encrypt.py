import io
import msoffcrypto
import openpyxl
import pandas as pd
import subprocess

# path1 = r'C:\Users\patryk.waraksa\Downloads\Challenge_Excel\Plik 1.xlsx'
# path2 = r'C:\Users\patryk.waraksa\Downloads\Challenge_Excel\Plik 2.xlsx'
# path3 = r'C:\Users\patryk.waraksa\Downloads\Challenge_Excel\Plik 3.xlsx'
# password1='plik1'
# password2='plik2'
# password3='plik3'
# zest_test_path = r"C:\Users\patryk.waraksa\Downloads\Challenge_Excel\zest_temp.xlsx"
# zest_pass = "zest"
# zest_encrypted_path = r"C:\Users\patryk.waraksa\Downloads\Challenge_Excel\zest_encrypted.xlsx"

path1 = input("path 1:")
path2 = input("path 2:")
path3 = input("path 3:")
password1= input("password 1:")
password2= input("password 2:")
password3= input("password 3:")
zest_test_path = input("zest_test_path:")
zest_pass = input("zest_pass:")
zest_encrypted_path = input("zest_encrypted_path:")

def decrypt_excel(path, password):
    decrypted_workbook = io.BytesIO()
    with open(path, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_workbook)
    return decrypted_workbook

def set_password(excel_file_path, pw):
    from pathlib import Path
    excel_file_path = Path(excel_file_path)
    vbs_script = \
    f"""' Save with password required upon opening

    Set excel_object = CreateObject("Excel.Application")
    Set workbook = excel_object.Workbooks.Open("{excel_file_path}")

    excel_object.DisplayAlerts = False
    excel_object.Visible = False

    workbook.SaveAs "{excel_file_path}",, "{pw}"

    excel_object.Application.Quit
    """
    # write
    vbs_script_path = excel_file_path.parent.joinpath("set_password.vbs")
    with open(vbs_script_path, "w") as file:
        file.write(vbs_script)
    #execute
    subprocess.call(['cscript.exe', str(vbs_script_path)])
    # remove
    # u mnie to nie działało więc wykomentowałem linijkę poniżej.
    # Rezultat - zapisuje się lokalnie kod VBA ;) sprobuj odkomentowac linijke ponizej i puscic skrypt.
    # vbs_script_path.unlink()
    return None

df1 = pd.read_excel(decrypt_excel(path1,password1))
df2 = pd.read_excel(decrypt_excel(path2,password2))
df3 = pd.read_excel(decrypt_excel(path3,password3))

vertical_concat = pd.concat([df1, df2, df3], axis=0)
print(vertical_concat)

vertical_concat.to_excel(zest_encrypted_path, index = False, header=True)
set_password(zest_encrypted_path, zest_pass)


