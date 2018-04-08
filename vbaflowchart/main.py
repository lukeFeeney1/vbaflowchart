"""
Program entry point.
"""

from win32com.client import Dispatch

def main():
    """Run the main process to generate the flow chart"""
    
    wbpath = 'C:\\Users\\Later\\Desktop\\Book1.xlsm'
    xl = Dispatch("Excel.Application")
    xl.Visible = 0
    xlwb = xl.Workbooks.Open(wbpath)

    module_list = []
    proc_list = []
    proc_string = ''

    #loop through each module and find procedures. Then loop through procedure to find other calls
    try:    
        for i in xlwb.VBProject.VBComponents:        
            code_mod = xlwb.VBProject.VBComponents(i.Name).CodeModule
            line_num = code_mod.CountOfDeclarationLines + 1
            count_of_lines = code_mod.CountOfLines

            #print(i.name)
            module_list.append(i.name)
            proc_string = ''
            
            while(line_num < count_of_lines):
                proc_name = code_mod.ProcOfLine(line_num, 0)            
                #print('> '+ proc_name[0])

                proc_string += proc_name[0] + ','

                curr_proc_start_line = code_mod.ProcStartLine(proc_name[0], 0)
                curr_proc_end_line = code_mod.ProcCountLines(proc_name[0], 0)
                line_num = curr_proc_start_line + curr_proc_end_line + 1

            proc_list.append(proc_string.split(",")[:-1])

    except Exception as e:
        print(e)

    finally:  
        print(dict(zip(module_list, proc_list))) 
        xlwb.Close(True)
        xl.Quit
        xl = None

if __name__ == '__main__':
    main()