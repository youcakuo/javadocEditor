#!/usr/bin/env python
# coding: utf-8

# In[1]:


# coding: utf-8

import re
import os
from git import Repo
from docx.api import Document
import win32com.client

VERSION = '1.3.2'


def get_spec_file(txnId):
    #print('[debug]get_spec_file')
    spec_path = "\\\\10.204.1.80\\tfb\\T-NBTS 專案_第二階段\\工作區\\個人工作區\\Shen\\javadoc\\javadoc_spec"
    all_spec_path_80='\\\\10.204.1.80\\tfb\\T-NBTS 專案_第二階段\\!IISI_FILE\\81第二階段-需求規格書'
    all_spec_path_89='\\\\172.16.240.89\\E_Disk\\!IISI_FILE\\81第二階段-需求規格書'
    result = ''
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.docx') and txnId in file:
                if not '歷史' in root and not '~$' in file:
                    #result = root + '\\' + file
                    result = os.path.join(root, file)
    if not result:
        for file in os.listdir(spec_path):
            if file.endswith(".docx") and txnId in file:
                #result = spec_path + '\\' + file
                result = os.path.join(spec_path, file)
    if not result:
        for root, dirs, files in os.walk(all_spec_path_89):
            for file in files:
                if file.endswith('.docx') and txnId in file:
                    if not '歷史' in root and not '~$' in file:
                        #result = root + '\\' + file
                        result = os.path.join(root, file)
    if not result:
        for root, dirs, files in os.walk(all_spec_path_80):
            for file in files:
                if file.endswith('.docx') and txnId in file:
                    if not '歷史' in root and not '~$' in file:
                        #result = root + '\\' + file
                        result = os.path.join(root, file)
    if not result:
        for root, dirs, files in os.walk(all_spec_path_89):
            for file in files:
                if file.endswith('.doc') and txnId in file:
                    if not '歷史' in root and not '~$' in file:
                        #result = root + '\\' + file
                        result = os.path.join(root, file)
    return result

def convert_doc_to_docx(file):
    if file.endswith('.doc'):
        doc = word.Documents.Open(file_path)
        docx_file = '{0}{1}'.format(file_path, 'x')
        doc.SaveAs(docx_file, 12)
        doc.Close()
        word.Quit()
        return docx_file
    else:
        return ''
        

def specParser(txnId):
    print('search spec for ' + txnId)
    spec_data1 = []
    spec_data2 = []
    spec_data3 = ''
    spec_data = []
    file = ''
    try:
        file = get_spec_file(txnId)
        print('[specParser] get_spec_file: ' + file)
    except:
        print('error when try to find sepcification')
    if not file:
        print('no valid spec found')
    else:
        converted_docx_file = convert_doc_to_docx(file)
        print('sepcification:' + file)
        if converted_docx_file:
            print('convert doc extension to docx')
        document = Document(converted_docx_file) if converted_docx_file else Document(file)
        print('spec open done')
        for table in document.tables:
            spec_data1 = []
            spec_data2 = []
            keys = None
            for i, row in enumerate(table.rows):
                text = (cell.text for cell in row.cells)
                if i == 0:
                    keys = tuple(text)
                    if keys == ('序號', '序號欄位名稱', 'I/O', '資料型態', '畫面元件', '格式化', '預設值', '必輸', '唯讀', '隱藏', '屬性及檢核'):
                        #print('欄位屬性')
                        continue
                    elif keys == ('序號', '欄位名稱', '處理方式'):
                        #print('欄位檢核說明')
                        continue
                    elif keys[0] == '交易初始化處理':
                        #print('交易初始化處理')
                        tmp = keys[1].replace('N/A','').rstrip()
                        if tmp:
                            spec_data3 = tmp
                        break
                    else:
                        break
                row_data = dict(zip(keys, text))
                if '屬性及檢核' in keys and row_data['屬性及檢核']:
                    row_data['屬性及檢核'] = row_data['屬性及檢核'].replace('\n\n',',').replace('\n',',').replace(',,',',').replace('；,',',')
                    spec_data1.append(row_data)
                elif '處理方式' in keys and row_data['處理方式']:
                    row_data['處理方式'] = row_data['處理方式'].replace('\n\n',',').replace('\n',',').replace(',,',',').replace('；,',',')
                    spec_data2.append(row_data)
            if spec_data1 or spec_data2:
                spec_data.append(spec_data1)
                spec_data.append(spec_data2)
        spec_data.append(spec_data3)
    return spec_data
            
def fileNameToTxnId(fileName):
    ftokens = re.split(r'[/]', fileName)
    return ftokens[len(ftokens)-1][2:10]

def getSpec1Content(spec):
    result = ''
    if not spec:
        return result
    size = len(spec)-1
    isMutiple = False
    if size > 2:
        isMutiple = True
        preTab = '     *   '
    else:
        preTab = '     * '
    for i in range(size):
        if i%2 == 1:
            continue
        rowdata = spec[i]
        if isMutiple:
            result = result + '     * 第'+str(int(i/2))+'畫面:' + '\n'
        for row in rowdata:
            if len(row['屬性及檢核']) > 20:
                result = result + preTab + row['序號欄位名稱'].replace('\n','') + ': ' + row['屬性及檢核'] + '\n'
    #print('spec1:' + result)
    return result
                
def getSpec2Content(spec):
    result = ''
    if not spec:
        return result
    size = len(spec)-1
    isMutiple = False
    if size > 2:
        isMutiple = True
        preTab = '     *   '
    else:
        preTab = '     * '
    for i in range(size):
        if i%2 == 0:
            continue
        rowdata = spec[i]
        if isMutiple:
            result = result + '     * 第'+str(int(i/2))+'畫面:' + '\n'
        for row in rowdata:
            if len(row['處理方式']) > 20:
                result = result + preTab + row['欄位名稱'].replace('\n','') + ': ' + row['處理方式'] + '\n'
    return result
                
def getSpec3Content(spec):
    result = ''
    if not spec:
        return result
    if spec[len(spec) -1]:
        result =  '     * ' + spec[len(spec)-1]
        result = result.replace('\n','\n     * ') + '\n'
    return result
    
def find_modified_source_files():
    path = '.'
    java_files = [f for f in os.listdir(path) if f.endswith('.java')]

    if not java_files:
        os.chdir("c://iisi/infinity-developer/repos/infinity-application-tfbnbts-transactions")
        repo = Repo('.')
        result = repo.git.status()
        tokens = re.split(r'[\n]', result)
        for token in tokens:
            if "modified:  " in token:
                filetokens = re.split(r'[ ]', token)
                fileToCheck = filetokens[len(filetokens) -1]
                if fileToCheck.endswith('.java'):
                    java_files.append(fileToCheck)  
    return java_files
    
def javadoc_template_for_sourcefile(javaFile):
    print('process file: ' + javaFile)
    try:
        spec_data = specParser(fileNameToTxnId(javaFile))
    except:
        spec_data = []
        print('Error when specParser')
    newlines = []   
    scriptTag = ''
    funcName = ''
    with open(javaFile, 'r', encoding='utf-8') as f:
        f_content = f.readlines()
        for i, line in enumerate(f_content):
            if '[method_name]' in str(line):
                scriptTag = ''
                for j in range(20):
                    if '*' in str(f_content[i+j]):
                        continue
                    line1 = str(f_content[i+j])
                    #line2 = str(f_content[i+j+1])
                    if 'private' in line1 or 'public' in line1 or 'protected' in line1:
                        scriptTag = 'CommentScriptlet'
                        #funcString = line1
                    else:
                        for n in range(5):                            
                            #funcString = line2
                            line1 = str(f_content[i+j+n])
                            tokens = re.split( r'[@(]', line1 )

                            #line1 = line2
                            if tokens[1] == 'CommentScriptlet' or tokens[1] == 'RelationshipScriptlet':
                                scriptTag = tokens[1]
                            line1 = str(f_content[i+j+n+1])
                            if not '@' in line1:
                                if scriptTag == '':
                                    scriptTag = 'CommentScriptlet'
                                break
                    tokens = re.split(r'[(]', line1)
                    line1 = tokens[0]
                    tokens = re.split(r'[ ]', line1)
                    funcName = tokens[len(tokens) -1]
                    break
                line = line.replace('[method_name]', '#' + scriptTag + ': ' + funcName)
                newlines.append(line)
            elif '[override_name]' in str(line):
                scriptTag = 'CommentScriptlet'
                for j in range(20):
                    if '*' in str(f_content[i+j]):
                        continue
                    line1 = str(f_content[i+j+1])
                    tokens = re.split(r'[(]', line1)
                    #print(tokens)
                    line1 = tokens[0]
                    tokens = re.split(r'[ ]', line1)
                    #print(tokens)
                    funcName = tokens[len(tokens)-1]
                    break
                line = line.replace('[override_name]', '#' + 'Method: ' + funcName)
                newlines.append(line)
                line = "     * #UsedByScriptlet: CrossValidation_Rule1\n"
                newlines.append(line)
            else:
                if '@param f\n' in line:
                    line = line.replace('@param f', '@param f 流程Facade')
                elif '@param n\n' in line:
                    line = line.replace('@param n', '@param n 通知')
                elif '@param cs\n' in line:
                    line = line.replace('@param cs', '@param cs 來源交易內文')
                elif '@param ct\n' in line:
                    line = line.replace('@param ct', '@param ct 目的交易內文')
                elif '@param c\n' in line:
                    line = line.replace('@param c', '@param c 交易內文')
                elif '@throws Throwable\n' in line:
                    line = line.replace('@throws Throwable', '@throws Throwable 例外錯誤')
                elif '@return\n' in line:
                    if scriptTag == 'RelationshipScriptlet':
                        line = line.replace('@return', '@return true為要執行關聯模組，false為不執行') 
                    elif funcName == 'InputController_1_FinishInputCondition':
                        line = line.replace('@return', '@return true為交易結束、false為尚未結束')
                    elif funcName == 'defaultBeforeInputConditions':
                        line = line.replace('@return', '@return true為可開放輸入、false為不可開放輸入')
                elif '[method_desc]' in line or '[override_desc]' in line:
                    if scriptTag == 'RelationshipScriptlet':
                        newlines.append('     * 關聯條件運算式\n')
                    elif funcName == 'doCrossValidationWhenAction':
                        newlines.append('     * 按鈕點擊的檢核與處理\n')
                    elif funcName == 'ActionControl':
                        newlines.append('     * 按鈕點擊的檢核與處理\n')
                    elif funcName == 'doCrossValidationWhenFieldInputCompleted':
                        newlines.append('     * 欄位輸入完畢時的檢核與處理\n')
                        if getSpec2Content(spec_data):
                            line = getSpec2Content(spec_data)
                    elif funcName == 'doCrossValidationWhenConfirmed':
                        newlines.append('     * 交易確認執行的檢核與處理\n')
                    elif funcName == 'FieldControl':
                        newlines.append('     * 欄位輸入完畢時的檢核與處理\n')
                        if getSpec1Content(spec_data):
                            line = getSpec1Content(spec_data)
                    elif funcName == 'FIELD_INPUT':
                        newlines.append('     * 欄位輸入完畢時的檢核與處理\n')
                        if getSpec1Content(spec_data):
                            line = getSpec1Content(spec_data)
                    elif 'ClientBeforeSendCBR003' in funcName:
                        newlines.append('     * 交易執行前的處理 (央媒查扣)\n')
                    elif 'ClientAfterSendCBR003' in funcName:
                        newlines.append('     * 交易執行後的處理 (央媒查扣)\n')
                    elif 'ClientBeforeSendCBR004' in funcName:
                        newlines.append('     * 交易執行前的處理 (央媒查扣沖正)\n')
                    elif 'ClientAfterSendCBR004' in funcName:
                        newlines.append('     * 交易執行後的處理 (央媒查扣沖正)\n')
                    elif 'ClientBeforeSend' in funcName:
                        newlines.append('     * 交易執行前的處理\n')
                    elif 'ClientAfterSend' in funcName:
                        newlines.append('     * 交易執行後的處理\n')
                    elif 'CBS' in funcName:
                        newlines.append('     * 交易執行前的處理\n')
                    elif 'CAS' in funcName:
                        newlines.append('     * 交易執行後的處理\n')
                    elif 'PatternInitial' in funcName:
                        newlines.append('     * 交易初始化\n')
                        if getSpec3Content(spec_data):
                            line = getSpec3Content(spec_data)
                    elif funcName == 'SetComposeTelegram':
                        newlines.append('     * 組合電文執行前的處理\n')
                    elif funcName == 'prepareCombineTelegram':
                        newlines.append('     * 組合電文執行前的處理\n')
                    elif funcName == 'defaultBeforeInputConditions':
                        newlines.append('     * 輸入模式開啟前的檢核與處理\n')
                    elif 'InputController' in funcName:
                        newlines.append('     * 交易確認執行完畢後，判斷是否應結束\n')
                        line = '     * 依據transactionState判斷交易是否結束\n'
                    funcName = ''
                    #scriptTag = ''
                newlines.append(line)

    with open(javaFile, 'w', encoding='utf-8') as f:
        for line in newlines:
            f.write(line)
        print('finish file: ' + javaFile + '\n')

    
def main():
    print('version: ' + VERSION)
    print('----------------------------------------------------------------------------')

    java_files = find_modified_source_files()
    print('modified files:')
    for mfiles in java_files:
        print('  ' + mfiles)
    print('\n----------------------------------------------------------------------------')
    for javaFile in java_files:
        javadoc_template_for_sourcefile(javaFile)
                
if __name__ == '__main__':
    try:
        main()
    finally:
        print('press Enter to continue...')
        input()

