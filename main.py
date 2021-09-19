
from kivymd.app import MDApp
from kivy.lang import Builder
import xlrd
from xlrd import sheet


KV = '''

MDScreen:
    MDTextField:
        id: quarto
        hint_text: "Digite o Quarto" 
        pos_hint: {"center_y": .5, 'center_x': .5} 
        size_hint_x: .40
        font_size: 60
        mode: 'fill' 
        halign: 'center'
        max_text_length: 4
        input_filter: 'int'
        input_type:'number'

    MDLabel:
        id: OK
        text:''
        halign: 'center'
        pos_hint: {"center_y": .1} 
        theme_text_color: "Custom"
        text_color: 'B71C1C'
        font_size: 70

   
    BoxLayout:
        orientation: 'vertical'
        MDToolbar:
            title: 'Localização de Antenas'
        estrutura:
       
    
<estrutura@FloatLayout>:

    Image:
        source: 'Ibero.png'
        size: self.texture_size
        pos_hint: {"center_y": .8, 'center_x': .5} 

    
    MDRaisedButton:
        text: "BUSCAR"
        on_release: app.consulta()
        font_size: "30sp"
        pos_hint: {"center_y": .3, 'center_x': .5} 

'''

class MyApp(MDApp):
    def build(self):
        return Builder.load_string(KV)
       
    def consulta(self):
        antena1 = self.root.ids.quarto.text
        antena = int(antena1)
        wb = xlrd.open_workbook('Consulta_Antenas.xls')
        ws = wb.sheet_by_index(0)
        sucesso = ''
    
 
        for sheet in wb.sheets():
            for linha in range(sheet.nrows):
                row = sheet.row(linha)
                for coluna, cell in enumerate(row):
                    if  cell.value == antena:
                        sucesso = True
                        ws = wb.sheet_by_index(0)
                        porta = ws.cell_value(rowx=linha, colx=1)
                        switch = ws.cell_value(rowx=linha, colx=2)
                   
                                  
        if sucesso == True:
            self.root.ids.OK.text = (f' A antena {antena} está na porta {porta} do {switch}')
        else:
            self.root.ids.OK.text = 'Este quarto não existe!'    
                      
MyApp().run()

