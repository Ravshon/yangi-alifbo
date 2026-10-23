import os
import sys
import shutil
import zipfile
import subprocess
import win32com.client
import tempfile

BASE_DIR = r"C:\Users\Ravshon\Desktop\YangiAlifbo"
DOTM_PATH = os.path.join(BASE_DIR, "YangiAlifbo.dotm")
KLAV_DIR = os.path.join(BASE_DIR, "YangiAlifbo_Klaviatura")
DOTM_KLAV = os.path.join(KLAV_DIR, "YangiAlifbo.dotm")

STARTUP_DIR = os.path.expandvars(r"%APPDATA%\Microsoft\Word\STARTUP")
DOTM_STARTUP = os.path.join(STARTUP_DIR, "YangiAlifbo.dotm")

# 1. Word jarayonlarini xavfsiz to'xtatish
subprocess.run(["taskkill", "/f", "/im", "winword.exe"], capture_output=True)

print("--- 1-qadam: YangiAlifbo.dotm shablonini yaratish ---")
word = win32com.client.Dispatch('Word.Application')
word.Visible = False
doc = word.Documents.Add()

vba_master = '''Option Explicit

Private Sub RepText(rng As Range, ByVal findWhat As String, ByVal replaceWith As String, Optional ByVal matchCase As Boolean = True)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findWhat
        .Replacement.Text = replaceWith
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .matchCase = matchCase
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Function GetWorkingRange() As Range
    If ActiveDocument Is Nothing Then
        Set GetWorkingRange = Nothing
        Exit Function
    End If
    If Selection.Type = wdSelectionIP Or Selection.Range.Characters.Count <= 1 Then
        Set GetWorkingRange = ActiveDocument.Content
    Else
        Set GetWorkingRange = Selection.Range
    End If
End Function

' ========================================================
' 1. LOTIN -> YANGI LOTIN
' ========================================================
Public Sub LotinToYangiLotin(Optional control As Object = Nothing)
    Dim rng As Range
    Set rng = GetWorkingRange()
    If rng Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Call RepText(rng, "SH", ChrW(350), True)
    Call RepText(rng, "Sh", ChrW(350), True)
    Call RepText(rng, "sh", ChrW(351), True)
    Call RepText(rng, "CH", ChrW(199), True)
    Call RepText(rng, "Ch", ChrW(199), True)
    Call RepText(rng, "ch", ChrW(231), True)
    
    Dim apos As Variant
    apos = Array("'", ChrW(8216), ChrW(8217), ChrW(699), ChrW(700), "`", ChrW(180))
    Dim i As Long
    For i = LBound(apos) To UBound(apos)
        Call RepText(rng, "O" & CStr(apos(i)), ChrW(214), True)
        Call RepText(rng, "o" & CStr(apos(i)), ChrW(246), True)
        Call RepText(rng, "G" & CStr(apos(i)), ChrW(286), True)
        Call RepText(rng, "g" & CStr(apos(i)), ChrW(287), True)
    Next i
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    Application.StatusBar = "Matn yangi lotin alifbosiga muvaffaqiyatli o'tkazildi! (Bekor qilish: Ctrl+Z)"
End Sub

' ========================================================
' 2. YANGI LOTIN -> ESKI LOTIN
' ========================================================
Public Sub YangiLotinToEskiLotin(Optional control As Object = Nothing)
    Dim rng As Range
    Set rng = GetWorkingRange()
    If rng Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Call RepText(rng, ChrW(350), "Sh", True)
    Call RepText(rng, ChrW(351), "sh", True)
    Call RepText(rng, ChrW(199), "Ch", True)
    Call RepText(rng, ChrW(231), "ch", True)
    
    Call RepText(rng, ChrW(214), "O'", True)
    Call RepText(rng, ChrW(246), "o'", True)
    Call RepText(rng, ChrW(286), "G'", True)
    Call RepText(rng, ChrW(287), "g'", True)
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    Application.StatusBar = "Matn eski lotin alifbosiga muvaffaqiyatli o'tkazildi! (Bekor qilish: Ctrl+Z)"
End Sub

' ========================================================
' 3. KIRILL -> YANGI LOTIN
' ========================================================
Public Sub KirillToYangiLotin(Optional control As Object = Nothing)
    Dim rng As Range
    Set rng = GetWorkingRange()
    If rng Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Call DoKirillToLotin(rng, True)
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    Application.StatusBar = "Matn Kirilldan Yangi Lotinga muvaffaqiyatli o'tkazildi! (Bekor qilish: Ctrl+Z)"
End Sub

' ========================================================
' 4. KIRILL -> ESKI LOTIN
' ========================================================
Public Sub KirillToEskiLotin(Optional control As Object = Nothing)
    Dim rng As Range
    Set rng = GetWorkingRange()
    If rng Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Call DoKirillToLotin(rng, False)
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    Application.StatusBar = "Matn Kirilldan Eski Lotinga muvaffaqiyatli o'tkazildi! (Bekor qilish: Ctrl+Z)"
End Sub

' ========================================================
' 5. LOTIN -> KIRILL
' ========================================================
Public Sub LotinToKirill(Optional control As Object = Nothing)
    Dim rng As Range
    Set rng = GetWorkingRange()
    If rng Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Call DoLotinToKirill(rng)
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    Application.StatusBar = "Matn Lotindan Kirill alifbosiga muvaffaqiyatli o'tkazildi! (Bekor qilish: Ctrl+Z)"
End Sub

' Yordamchi: Kirill -> Lotin
Private Sub DoKirillToLotin(rng As Range, isNew As Boolean)
    If isNew Then
        Call RepText(rng, ChrW(1064), ChrW(350), True)
        Call RepText(rng, ChrW(1096), ChrW(351), True)
        Call RepText(rng, ChrW(1065), ChrW(350), True)
        Call RepText(rng, ChrW(1097), ChrW(351), True)
        Call RepText(rng, ChrW(1063), ChrW(199), True)
        Call RepText(rng, ChrW(1095), ChrW(231), True)
        Call RepText(rng, ChrW(1038), ChrW(214), True)
        Call RepText(rng, ChrW(1118), ChrW(246), True)
        Call RepText(rng, ChrW(1170), ChrW(286), True)
        Call RepText(rng, ChrW(1171), ChrW(287), True)
    Else
        Call RepText(rng, ChrW(1064), "Sh", True)
        Call RepText(rng, ChrW(1096), "sh", True)
        Call RepText(rng, ChrW(1065), "Sh", True)
        Call RepText(rng, ChrW(1097), "sh", True)
        Call RepText(rng, ChrW(1063), "Ch", True)
        Call RepText(rng, ChrW(1095), "ch", True)
        Call RepText(rng, ChrW(1038), "O'", True)
        Call RepText(rng, ChrW(1118), "o'", True)
        Call RepText(rng, ChrW(1170), "G'", True)
        Call RepText(rng, ChrW(1171), "g'", True)
    End If
    
    Call RepText(rng, ChrW(1025), "Yo", True)
    Call RepText(rng, ChrW(1105), "yo", True)
    Call RepText(rng, ChrW(1070), "Yu", True)
    Call RepText(rng, ChrW(1102), "yu", True)
    Call RepText(rng, ChrW(1071), "Ya", True)
    Call RepText(rng, ChrW(1103), "ya", True)
    Call RepText(rng, ChrW(1062), "Ts", True)
    Call RepText(rng, ChrW(1094), "ts", True)
    
    Dim vChars As Variant
    vChars = Array(ChrW(1040), ChrW(1072), ChrW(1054), ChrW(1086), ChrW(1059), ChrW(1091), ChrW(1048), ChrW(1080), ChrW(1069), ChrW(1101), ChrW(1045), ChrW(1077), ChrW(1025), ChrW(1105), ChrW(1070), ChrW(1102), ChrW(1071), ChrW(1103), ChrW(1038), ChrW(1118), ChrW(1098), ChrW(1066), " ", Chr(34), "'", "(", "-", vbTab)
    Dim k As Long
    For k = LBound(vChars) To UBound(vChars)
        Call RepText(rng, CStr(vChars(k)) & ChrW(1045), CStr(vChars(k)) & "Ye", True)
        Call RepText(rng, CStr(vChars(k)) & ChrW(1077), CStr(vChars(k)) & "ye", True)
    Next k
    
    If Len(rng.Text) > 0 Then
        If Left(rng.Text, 1) = ChrW(1045) Then rng.Characters(1).Text = "Ye"
        If Left(rng.Text, 1) = ChrW(1077) Then rng.Characters(1).Text = "ye"
    End If
    
    Call RepText(rng, ChrW(1045), "E", True)
    Call RepText(rng, ChrW(1077), "e", True)
    
    Call RepText(rng, ChrW(1040), "A", True)
    Call RepText(rng, ChrW(1072), "a", True)
    Call RepText(rng, ChrW(1041), "B", True)
    Call RepText(rng, ChrW(1073), "b", True)
    Call RepText(rng, ChrW(1042), "V", True)
    Call RepText(rng, ChrW(1074), "v", True)
    Call RepText(rng, ChrW(1043), "G", True)
    Call RepText(rng, ChrW(1075), "g", True)
    Call RepText(rng, ChrW(1044), "D", True)
    Call RepText(rng, ChrW(1076), "d", True)
    Call RepText(rng, ChrW(1046), "J", True)
    Call RepText(rng, ChrW(1078), "j", True)
    Call RepText(rng, ChrW(1047), "Z", True)
    Call RepText(rng, ChrW(1079), "z", True)
    Call RepText(rng, ChrW(1048), "I", True)
    Call RepText(rng, ChrW(1080), "i", True)
    Call RepText(rng, ChrW(1049), "Y", True)
    Call RepText(rng, ChrW(1081), "y", True)
    Call RepText(rng, ChrW(1050), "K", True)
    Call RepText(rng, ChrW(1082), "k", True)
    Call RepText(rng, ChrW(1051), "L", True)
    Call RepText(rng, ChrW(1083), "l", True)
    Call RepText(rng, ChrW(1052), "M", True)
    Call RepText(rng, ChrW(1084), "m", True)
    Call RepText(rng, ChrW(1053), "N", True)
    Call RepText(rng, ChrW(1085), "n", True)
    Call RepText(rng, ChrW(1054), "O", True)
    Call RepText(rng, ChrW(1086), "o", True)
    Call RepText(rng, ChrW(1055), "P", True)
    Call RepText(rng, ChrW(1087), "p", True)
    Call RepText(rng, ChrW(1056), "R", True)
    Call RepText(rng, ChrW(1088), "r", True)
    Call RepText(rng, ChrW(1057), "S", True)
    Call RepText(rng, ChrW(1089), "s", True)
    Call RepText(rng, ChrW(1058), "T", True)
    Call RepText(rng, ChrW(1090), "t", True)
    Call RepText(rng, ChrW(1059), "U", True)
    Call RepText(rng, ChrW(1091), "u", True)
    Call RepText(rng, ChrW(1060), "F", True)
    Call RepText(rng, ChrW(1092), "f", True)
    Call RepText(rng, ChrW(1061), "X", True)
    Call RepText(rng, ChrW(1093), "x", True)
    Call RepText(rng, ChrW(1069), "E", True)
    Call RepText(rng, ChrW(1101), "e", True)
    Call RepText(rng, ChrW(1178), "Q", True)
    Call RepText(rng, ChrW(1179), "q", True)
    Call RepText(rng, ChrW(1120), "H", True)
    Call RepText(rng, ChrW(1121), "h", True)
    Call RepText(rng, ChrW(1066), "'", True)
    Call RepText(rng, ChrW(1098), "'", True)
    Call RepText(rng, ChrW(1068), "", True)
    Call RepText(rng, ChrW(1100), "", True)
End Sub

' Yordamchi: Lotin -> Kirill
Private Sub DoLotinToKirill(rng As Range)
    Dim apos As Variant
    apos = Array("'", ChrW(8216), ChrW(8217), ChrW(699), ChrW(700), "`", ChrW(180))
    Dim i As Long
    For i = LBound(apos) To UBound(apos)
        Call RepText(rng, "O" & CStr(apos(i)), ChrW(1038), True)
        Call RepText(rng, "o" & CStr(apos(i)), ChrW(1118), True)
        Call RepText(rng, "G" & CStr(apos(i)), ChrW(1170), True)
        Call RepText(rng, "g" & CStr(apos(i)), ChrW(1171), True)
    Next i
    
    Call RepText(rng, ChrW(214), ChrW(1038), True)
    Call RepText(rng, ChrW(246), ChrW(1118), True)
    Call RepText(rng, ChrW(286), ChrW(1170), True)
    Call RepText(rng, ChrW(287), ChrW(1171), True)
    Call RepText(rng, ChrW(350), ChrW(1064), True)
    Call RepText(rng, ChrW(351), ChrW(1096), True)
    Call RepText(rng, ChrW(199), ChrW(1063), True)
    Call RepText(rng, ChrW(231), ChrW(1095), True)
    
    Call RepText(rng, "SH", ChrW(1064), True)
    Call RepText(rng, "Sh", ChrW(1064), True)
    Call RepText(rng, "sh", ChrW(1096), True)
    Call RepText(rng, "CH", ChrW(1063), True)
    Call RepText(rng, "Ch", ChrW(1063), True)
    Call RepText(rng, "ch", ChrW(1095), True)
    Call RepText(rng, "YO", ChrW(1025), True)
    Call RepText(rng, "Yo", ChrW(1025), True)
    Call RepText(rng, "yo", ChrW(1105), True)
    Call RepText(rng, "YU", ChrW(1070), True)
    Call RepText(rng, "Yu", ChrW(1070), True)
    Call RepText(rng, "yu", ChrW(1102), True)
    Call RepText(rng, "YA", ChrW(1071), True)
    Call RepText(rng, "Ya", ChrW(1071), True)
    Call RepText(rng, "ya", ChrW(1103), True)
    Call RepText(rng, "YE", ChrW(1045), True)
    Call RepText(rng, "Ye", ChrW(1045), True)
    Call RepText(rng, "ye", ChrW(1077), True)
    Call RepText(rng, "TS", ChrW(1062), True)
    Call RepText(rng, "Ts", ChrW(1062), True)
    Call RepText(rng, "ts", ChrW(1094), True)
    
    Dim bChars As Variant
    bChars = Array(" ", Chr(34), "(", "-", ChrW(171), "[", vbTab)
    Dim b As Long
    For b = LBound(bChars) To UBound(bChars)
        Call RepText(rng, CStr(bChars(b)) & "E", CStr(bChars(b)) & ChrW(1069), True)
        Call RepText(rng, CStr(bChars(b)) & "e", CStr(bChars(b)) & ChrW(1101), True)
    Next b
    
    If Len(rng.Text) > 0 Then
        If Left(rng.Text, 1) = "E" Then rng.Characters(1).Text = ChrW(1069)
        If Left(rng.Text, 1) = "e" Then rng.Characters(1).Text = ChrW(1101)
    End If
    
    Call RepText(rng, "A", ChrW(1040), True)
    Call RepText(rng, "a", ChrW(1072), True)
    Call RepText(rng, "B", ChrW(1041), True)
    Call RepText(rng, "b", ChrW(1073), True)
    Call RepText(rng, "D", ChrW(1044), True)
    Call RepText(rng, "d", ChrW(1076), True)
    Call RepText(rng, "E", ChrW(1045), True)
    Call RepText(rng, "e", ChrW(1077), True)
    Call RepText(rng, "F", ChrW(1060), True)
    Call RepText(rng, "f", ChrW(1092), True)
    Call RepText(rng, "G", ChrW(1043), True)
    Call RepText(rng, "g", ChrW(1075), True)
    Call RepText(rng, "H", ChrW(1120), True)
    Call RepText(rng, "h", ChrW(1121), True)
    Call RepText(rng, "I", ChrW(1048), True)
    Call RepText(rng, "i", ChrW(1080), True)
    Call RepText(rng, "J", ChrW(1046), True)
    Call RepText(rng, "j", ChrW(1078), True)
    Call RepText(rng, "K", ChrW(1050), True)
    Call RepText(rng, "k", ChrW(1082), True)
    Call RepText(rng, "L", ChrW(1051), True)
    Call RepText(rng, "l", ChrW(1083), True)
    Call RepText(rng, "M", ChrW(1052), True)
    Call RepText(rng, "m", ChrW(1084), True)
    Call RepText(rng, "N", ChrW(1053), True)
    Call RepText(rng, "n", ChrW(1085), True)
    Call RepText(rng, "O", ChrW(1054), True)
    Call RepText(rng, "o", ChrW(1086), True)
    Call RepText(rng, "P", ChrW(1055), True)
    Call RepText(rng, "p", ChrW(1087), True)
    Call RepText(rng, "Q", ChrW(1178), True)
    Call RepText(rng, "q", ChrW(1179), True)
    Call RepText(rng, "R", ChrW(1056), True)
    Call RepText(rng, "r", ChrW(1088), True)
    Call RepText(rng, "S", ChrW(1057), True)
    Call RepText(rng, "s", ChrW(1089), True)
    Call RepText(rng, "T", ChrW(1058), True)
    Call RepText(rng, "t", ChrW(1090), True)
    Call RepText(rng, "U", ChrW(1059), True)
    Call RepText(rng, "u", ChrW(1091), True)
    Call RepText(rng, "V", ChrW(1042), True)
    Call RepText(rng, "v", ChrW(1074), True)
    Call RepText(rng, "X", ChrW(1061), True)
    Call RepText(rng, "x", ChrW(1093), True)
    Call RepText(rng, "Y", ChrW(1049), True)
    Call RepText(rng, "y", ChrW(1081), True)
    Call RepText(rng, "Z", ChrW(1047), True)
    Call RepText(rng, "z", ChrW(1079), True)
    
    For i = LBound(apos) To UBound(apos)
        Call RepText(rng, CStr(apos(i)), ChrW(1098), True)
    Next i
End Sub
'''

vb_comp = doc.VBProject.VBComponents.Add(1)
vb_comp.Name = "YangiAlifbo"
vb_comp.CodeModule.AddFromString(vba_master)

if os.path.exists(DOTM_PATH):
    os.remove(DOTM_PATH)

doc.SaveAs2(DOTM_PATH, FileFormat=15)
doc.Close(False)
word.Quit()
print("Asosiy YangiAlifbo.dotm shabloni saqlandi.")

custom_ui_xml = '''<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabYangiAlifbo" label="Yangi Alifbo" insertAfterMso="TabView">
        
        <group id="grpLotin" label="Lotin Alifbosi">
          <button id="btnLotinToYangi"
                  label="Yangi alifboga o'girish"
                  size="large"
                  imageMso="SignatureLineInsert"
                  onAction="YangiAlifbo.LotinToYangiLotin"
                  screentip="Eski Lotin -> Yangi Lotin"
                  supertip="Hujjatdagi SH, CH, O', G' harflarini yangi alifboga (Ş, Ç, Ö, Ğ) aylantiradi." />
          <button id="btnYangiToEski"
                  label="Eski alifboga o'girish"
                  size="large"
                  imageMso="Undo"
                  onAction="YangiAlifbo.YangiLotinToEskiLotin"
                  screentip="Yangi Lotin -> Eski Lotin"
                  supertip="Yangi alifbodagi (Ş, Ç, Ö, Ğ) harflarni eski lotin alifbosiga (Sh, Ch, O', G') qaytaradi." />
        </group>

        <group id="grpKirillToLotin" label="Kirilldan Lotinga">
          <button id="btnKirillToYangi"
                  label="Kirill -> Yangi Lotin"
                  size="large"
                  imageMso="Translate"
                  onAction="YangiAlifbo.KirillToYangiLotin"
                  screentip="Kirill -> Yangi Lotin (Ş, Ç, Ö, Ğ)"
                  supertip="Kirill alifbosidagi matnni to'g'ridan-to'g'ri yangi lotin alifbosiga o'giradi." />
          <button id="btnKirillToEski"
                  label="Kirill -> Eski Lotin"
                  size="large"
                  imageMso="Spelling"
                  onAction="YangiAlifbo.KirillToEskiLotin"
                  screentip="Kirill -> Eski Lotin (Sh, Ch, O', G')"
                  supertip="Kirill alifbosidagi matnni eski lotin alifbosiga o'giradi." />
        </group>

        <group id="grpLotinToKirill" label="Lotindan Kirillga">
          <button id="btnLotinToKirill"
                  label="Lotindan Kirillga"
                  size="large"
                  imageMso="ReplaceDialog"
                  onAction="YangiAlifbo.LotinToKirill"
                  screentip="Lotin -> Kirill"
                  supertip="Har qanday lotin matnini (xoh yangi Ş, Ç, Ö, Ğ bo'lsin, xoh eski Sh, Ch, O', G' bo'lsin) to'g'ri Kirill alifbosiga o'giradi." />
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>'''

temp_dir = tempfile.mkdtemp()
try:
    with zipfile.ZipFile(DOTM_PATH, 'r') as zip_read:
        zip_read.extractall(temp_dir)
        
    customui_folder = os.path.join(temp_dir, 'customUI')
    os.makedirs(customui_folder, exist_ok=True)
    with open(os.path.join(customui_folder, 'customUI14.xml'), 'w', encoding='utf-8') as f:
        f.write(custom_ui_xml)
        
    rels_path = os.path.join(temp_dir, '_rels', '.rels')
    with open(rels_path, 'r', encoding='utf-8') as f:
        rels_content = f.read()
        
    rel_customui = '<Relationship Id="rIdCustomUI14" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
    if 'customUI14.xml' not in rels_content:
        rels_content = rels_content.replace('</Relationships>', f'  {rel_customui}\n</Relationships>')
        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_content)
            
    with zipfile.ZipFile(DOTM_PATH, 'w', zipfile.ZIP_DEFLATED) as zip_write:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, temp_dir)
                zip_write.write(abs_path, rel_path)
    print("YangiAlifbo.dotm to'liq qayta yaratildi.")
finally:
    shutil.rmtree(temp_dir)

# 2. YangiAlifbo_Klaviatura va Word STARTUP papkalariga nusxalash
shutil.copy2(DOTM_PATH, DOTM_KLAV)
print(f"Nusxalandi: {DOTM_KLAV}")

try:
    shutil.copy2(DOTM_PATH, DOTM_STARTUP)
    print(f"Nusxalandi: {DOTM_STARTUP}")
except Exception as e:
    print(f"STARTUP ga nusxalashda xatolik: {e}")

# 3. Word_YangiAlifbo_Addin.zip yaratish
addin_zip = os.path.join(BASE_DIR, "Word_YangiAlifbo_Addin.zip")
if os.path.exists(addin_zip): os.remove(addin_zip)
with zipfile.ZipFile(addin_zip, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(DOTM_PATH, "YangiAlifbo.dotm")
    z.write(os.path.join(BASE_DIR, "Word_Qoshimchasini_Ornatish.bat"), "Word_Qoshimchasini_Ornatish.bat")
    z.write(os.path.join(BASE_DIR, "Word_Qoshimchasini_Ochirish.bat"), "Word_Qoshimchasini_Ochirish.bat")
print("Word_YangiAlifbo_Addin.zip yaratildi.")

# 4. Yangi_Alifbo_Klaviatura.zip yaratish
klav_zip = os.path.join(BASE_DIR, "Yangi_Alifbo_Klaviatura.zip")
if os.path.exists(klav_zip): os.remove(klav_zip)
with zipfile.ZipFile(klav_zip, 'w', zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(KLAV_DIR):
        for file in files:
            abs_p = os.path.join(root, file)
            rel_p = os.path.relpath(abs_p, BASE_DIR)
            z.write(abs_p, rel_p)
print("Yangi_Alifbo_Klaviatura.zip yaratildi.")

# 5. YangiAlifbo_Setup.exe dasturini kompyutatsiyalash
csc_path = r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
cmd = [
    csc_path,
    "/target:winexe",
    f"/out:{os.path.join(BASE_DIR, 'YangiAlifbo_Setup.exe')}",
    f"/resource:{klav_zip},payload.zip",
    "/r:System.Windows.Forms.dll,System.Drawing.dll,System.IO.Compression.dll,System.IO.Compression.FileSystem.dll",
    os.path.join(BASE_DIR, "YangiAlifbo_Setup.cs")
]
res = subprocess.run(cmd, capture_output=True, text=True)
if res.returncode == 0:
    print("YangiAlifbo_Setup.exe muvaffaqiyatli kompilyatsiya qilindi!")
else:
    print("CSC kompilyatsiya xatosi:")
    print(res.stderr)
    sys.exit(1)

print("BARCHA ISHLAR MUVAFFAQIYATLI YAKUNLANDI!")
