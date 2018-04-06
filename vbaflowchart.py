from win32com.client import Dispatch
wbpath = 'C:\\Users\\Later\\Desktop\\Book1.xlsm'
xl = Dispatch("Excel.Application")
xl.Visible = 0
xlwb = xl.Workbooks.Open(wbpath)

# ITERATE THROUGH EACH VB COMPONENT (CLASS MODULE, STANDARD MODULE, USER FORMS)
try:    
    for i in xlwb.VBProject.VBComponents:        
        code_mod = xlwb.VBProject.VBComponents(i.Name).CodeModule
        #proc_kind = xlwb.VBProject.VBComponents.vbext_ProcKind
        line_num = code_mod.CountOfDeclarationLines + 1
        count_of_lines = code_mod.CountOfLines
        print('----------------')
        print(i.name)
        print('----------------')
        while(line_num < count_of_lines):
            proc_name = code_mod.ProcOfLine(line_num, 0)            
            print (proc_name[0])
            line_num = code_mod.ProcStartLine(proc_name[0], 0) + code_mod.ProcCountLines(proc_name[0], 0) + 1
except Exception as e:
    print(e)

finally:    
    # CLOSE AND SAVE AND UNINITIALIZE APP
    xlwb.Close(True)
    xl.Quit

    xl = None
