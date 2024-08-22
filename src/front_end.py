from collections.abc import Callable, Sequence
from typing import Any, Optional, Tuple, Union
import customtkinter as ctk
from tkinter import *
from PIL import ImageTk,Image
from tkinter import messagebox

from src.p_societaria_algoritm import PropostaSocietaria

from tkinter import ttk


from tkinter import filedialog

from typing import Dict

from docx import Document

from datetime import datetime

from typing import List

class Janela(ctk.CTk):

    def __init__(self, fg_color: str | Tuple[str] | None = None, **kwargs):
        super().__init__(fg_color, **kwargs)

        ctk.set_appearance_mode("Dark")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("blue")
        ctk.deactivate_automatic_dpi_awareness()

        self.title('Automação Proposta Societaria')
        self.geometry('1024x600')
        self.resizable(False,False)




class MainFrame(ctk.CTkFrame):

    def __init__(self, master: Any, width: int = 1024, height: int = 600, corner_radius: int | str | None = None, border_width: int | str | None = None, bg_color: str | Tuple[str] = "transparent", fg_color: str | Tuple[str] | None = None, border_color: str | Tuple[str] | None = None, background_corner_colors: Tuple[str | Tuple[str]] | None = None, overwrite_preferred_drawing_method: str | None = None, **kwargs):
        super().__init__(master, width, height, corner_radius, border_width, bg_color, fg_color, border_color, background_corner_colors, overwrite_preferred_drawing_method, **kwargs)

        self.lista_serviços = [
            'Constituição de empresa prestadora de Serviços',
            'Constituição de empresa comercial',
            'Alteração contratual',
            'Baixa de empresa - distrato social',
            'Obtenção certidões',
            'Regularização de empresa',
            'Atestado de residencia fiscal',
        ]

        self.honorarios_table = ttk.Treeview(self, columns=('Serviços'), show='headings',height=10)
        self.honorarios_table.column('Serviços', width=300)
        self.honorarios_table.heading('Serviços', text='Serviços')

        self.honorarios_table.place(x = 600, y = 160)

        self.honorarios_table_scrollbar = ctk.CTkScrollbar(self, command=self.honorarios_table.yview, fg_color='#6D6A6A', height=224, button_color='#252323')
        self.honorarios_table_scrollbar.place(x = 900, y = 160)

        self.honorarios_table.tag_configure('nome', background='#3D3838',foreground='#fff')
        self.set_treeview_style()


        self.empresa_label = ctk.CTkLabel(master = self,
                                          width=80, 
                                           height=20, 
                                              text = 'Empresa:',
                                               font=('Inter', 19, 'bold'),
                                                text_color='#fff' )

        self.empresa_label.place(x = 20, y = 40)

        self.empresa_entry = ctk.CTkEntry(master = self, 
                                            width=250,
                                            height=40,
                                            placeholder_text='Empresa:',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )
        
        self.empresa_entry.place(x =150, y = 40)

        self.empresa_entry.bind('<Return>', self.empresa_bind_command)    
        
        self.nome_responsável_label = ctk.CTkLabel(master = self, 
                                                   width=80,
                                                   height=20, 
                                                   text ='Nome:',
                                                   font = ('Inter', 19, 'bold'),
                                                   text_color='#fff')
        
        self.nome_responsável_label.place(x = 20, y = 110)

        self.nome_responsavel_entry = ctk.CTkEntry(master = self, 
                                            width=250,
                                            height=40,
                                            placeholder_text='Nome responsavel:',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )
        
        self.nome_responsavel_entry.place(x =150, y = 110)

        self.nome_responsavel_entry.bind('<Return>', self.nome_responsavel_bind_command)



        self.valor_honorário_label = ctk.CTkLabel(master=self, 
                                            width=80, 
                                            height=20, 
                                            text = 'Valor:',
                                            font=('Inter', 19, 'bold'),
                                            text_color='#fff'
                                            )
        
        self.valor_honorário_label.place(x = 20, y = 180)

        self.valor_honorário_entry= ctk.CTkEntry(master = self, 
                                            width=250,
                                            height=40,
                                            placeholder_text='2.800,00 (dois mil e oitocentos reais)',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )
        
        self.valor_honorário_entry.place(x =150, y = 180)

        self.valor_honorário_entry.bind('<Return>', self.valor_honorario_bind_command)
       


        self.data_label = ctk.CTkLabel(master = self, 
                                       height=20,
                                        width=80, 
                                        text = 'Data:',
                                         font=('Inter', 19, 'bold'),
                                          text_color='#fff' )
        
        self.data_label.place(x = 20, y =250 )

        self.data_entry= ctk.CTkEntry(master = self, 
                                            width=250,
                                            height=40,
                                            placeholder_text='ex: 05 de Agosto de 2024',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )
        
        self.data_entry.place(x =150, y = 250)

        self.data_entry.bind('<Return>', self.data_bind_command)
        

        self.mes_ano_label = ctk.CTkLabel(master=self, 
                                          width=80, 
                                          height=20,
                                          text = 'Mes-Ano:',
                                          font = ( 'Inter', 19, 'bold'),
                                          text_color='#fff'
                                          )


        self.mes_ano_label.place(x = 20,  y= 320)

        self.mes_ano_entry= ctk.CTkEntry(master = self, 
                                            width=250,
                                            height=40,
                                            placeholder_text='ex: Agosto de 2024',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )
        
        self.mes_ano_entry.place(x =150, y = 320)

        self.mes_ano_entry.bind('<Return>', self.mes_ano_bind_command)

        self.codigo_label = ctk.CTkLabel(master=self, 
                                          width=80, 
                                          height=20,
                                          text = 'N proposta:',
                                          font = ( 'Inter', 19, 'bold'),
                                          text_color='#fff'
                                          )


        self.codigo_label.place(x = 20,  y= 390)

        self.codigo_entry= ctk.CTkEntry(master = self, 
                                            width=250,
                                            height=40,
                                            placeholder_text='ex: 1111',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )
        
        self.codigo_entry.place(x =150, y = 390)


        self.honorario_entry = ctk.CTkEntry(master = self, 
                                            width=250,
                                            height=40,
                                            placeholder_text='Serviço:',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )
        
        self.honorario_entry.place(x =600, y = 400)


        self.honorario_buttom = ctk.CTkButton(master = self, 
                                               width =40,
                                                height=40,
                                                 corner_radius=5,
                                                  text_color='#fff',
                                                   command=self.ok_treeview_command, 
                                                    text = 'OK',
                                                     font=('Inter', 19, 'bold'))
        
        self.honorario_buttom.place(x = 870, y = 400 )
        


        

        self.gerar_buttom = ctk.CTkButton(master = self, 
                                          width = 140,
                                          height=70,
                                          text = 'GERAR',
                                          font=('Inter', 19, 'bold'),
                                          corner_radius=5,
                                          command = self.gerar_proposta,

                                          )
        
        self.gerar_buttom.place(x = 820, y = 490)


        self.tipo_proposta_string_var = ctk.StringVar(value = 'Serviço')
        self.tipo_proposta = ctk.CTkOptionMenu(self,
                                               values = self.lista_serviços,
                                               variable = self.tipo_proposta_string_var,
                                               width = 300, 
                                               height = 40,
                                                )
        
        self.tipo_proposta.place(x =600, y =40 )

        self.tipo_proposta_outros = ctk.CTkEntry(master = self, 
                                            width=300,
                                            height=40,
                                            placeholder_text='Outros...',
                                            text_color='#fff',
                                            corner_radius=5,
                                            border_color='#fff'
                                            )

        self.tipo_proposta_outros.place(x = 600, y = 85)


    def gerar_proposta(self):
        
        lista_termos_proposta = []

        for termo in self.honorarios_table.get_children():
            lista_termos_proposta.append(self.honorarios_table.item(termo)['values'][0])

        nome_empresa:str = self.empresa_entry.get()
        mes_ano:str =   self.mes_ano_entry.get()
        codigo:str = self.codigo_entry.get()
        data:str = self.data_entry.get()
        nome_responsavel:str = self.nome_responsavel_entry.get()
        proposta:str = None
        valor_honorario:str = self.valor_honorário_entry.get()

        if len(self.tipo_proposta_outros.get()) > 0:
            proposta = self.tipo_proposta_outros.get()
        
        else:
            proposta = self.tipo_proposta_string_var.get()

        
        try:

            _path = filedialog.askdirectory()

            proposta = PropostaSocietaria(nome_empresa=nome_empresa,
                                          mes_ano=mes_ano,
                                          codigo=codigo,
                                          data=data,
                                          nome_responsavel=nome_responsavel,
                                          valor_honorario=valor_honorario,
                                          proposta=proposta,
                                          lista_termos_proposta=lista_termos_proposta)
            

            proposta.gerar_proposta_societaria(path = _path)

            messagebox.showinfo(title='Documento Gerado', message = 'Documento gerado com sucesso!')

            self.clear_all()
            

        except Exception  as exception:

            messagebox.showerror(title= 'Exception', message=exception)

            raise exception


    def clear_all(self):

        self.clear_treeview()

        self.empresa_entry.delete(0,END)

        self.nome_responsavel_entry.delete(0, END)

        self.valor_honorário_entry.delete(0, END)

        self.codigo_entry.delete(0, END)

        self.data_entry.delete(0, END)

        self.mes_ano_entry.delete(0, END)

        self.tipo_proposta_outros.delete(0, END)

    def set_treeview_style(self):
        style = ttk.Style()
        style.theme_use("default")  

        style.configure("Treeview.Heading", foreground="white")
        style.configure("Treeview.Heading", background="#212D40" )




    def empresa_bind_command(self,e):

        self.nome_responsavel_entry.focus()

    
    def nome_responsavel_bind_command(self,e):

        self.valor_honorário_entry.focus()

    
    def valor_honorario_bind_command(self,e):

        self.data_entry.focus()

    
    def data_bind_command(self,e):

        self.mes_ano_entry.focus()

    def mes_ano_bind_command(self,e):

        self.codigo_entry.focus()



    def clear_treeview(self):

        for item in self.honorarios_table.get_children():

            self.honorarios_table.delete(item)

    

    def ok_treeview_command(self):

        honorario = [self.honorario_entry.get()]

        self.honorarios_table.insert('',  'end', values=honorario)

        self.honorario_entry.delete(0, END)


