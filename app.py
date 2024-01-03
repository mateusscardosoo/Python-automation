import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('./assets/produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(-1200,177, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    

    descricao_produto = linha[1].value
    pyperclip.copy(descricao_produto)
    pyautogui.click(-1200,256, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    

    categoria_produto = linha[2].value
    pyperclip.copy(categoria_produto)
    pyautogui.click(-1200,397, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(-1200,481, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    
    peso_produto = linha[4].value
    pyperclip.copy(peso_produto)
    pyautogui.click(-1200,569, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    dimensoes_produto = linha[5].value
    pyperclip.copy(dimensoes_produto)
    pyautogui.click(-1200,609, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    #Proxima página
    pyautogui.click(-1200,703, duration=1)
    sleep(3)

    preco_produto = linha[6].value
    pyperclip.copy(preco_produto)
    pyautogui.click(-1200,204, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    qtde_estoque_produto = linha[7].value
    pyperclip.copy(qtde_estoque_produto)
    pyautogui.click(-1200,286, duration=1)
    pyautogui.hotkey('ctrl', 'v')



    validade_produto = linha[8].value
    pyperclip.copy(validade_produto)
    pyautogui.click(-1200,371, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    cor_produto = linha[9].value
    pyperclip.copy(cor_produto)
    pyautogui.click(-1200,457, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    tamanho_produto = linha[10].value
    pyautogui.click(-1200,546, duration=1)

    if tamanho_produto == 'Pequeno':
        pyautogui.click(-1200,569, duration=1)


    elif tamanho_produto == 'Médio':
        pyautogui.click(-1200,592, duration=1)
 

    else:
        pyautogui.click(-1200,620, duration=1)

    
  
    material_produto = linha[11].value
    pyperclip.copy(material_produto)
    pyautogui.click(-1231,627, duration=1)
    pyautogui.hotkey('ctrl', 'v')

      #Proxima página
    pyautogui.click(-1218,693, duration=1)
    sleep(3)


    fabricante_produto = linha[12].value
    pyperclip.copy(fabricante_produto)
    pyautogui.click(-1226,219, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    origem_produto = linha[13].value
    pyperclip.copy(origem_produto)
    pyautogui.click(-1231,307, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    observacoes_produto = linha[14].value
    pyperclip.copy(observacoes_produto)
    pyautogui.click(-1242,384, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    gtin_produto = linha[15].value
    pyperclip.copy(gtin_produto)
    pyautogui.click(-1236,527, duration=1)
    pyautogui.hotkey('ctrl', 'v')


    prateleira_produto = linha[16].value
    pyperclip.copy(prateleira_produto)
    pyautogui.click(-1235,615, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Concluir
    pyautogui.click(-1225,673,duration=1)
    pyautogui.click(-515,187,duration=1)
    pyautogui.click(-730,449,duration=1)



