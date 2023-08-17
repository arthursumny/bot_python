# Esta é a versão 4.1 do Bot.unoesc - Envio de Mensagens Automáticas

import pandas as pd
import re
import pywhatkit as kit
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import Radiobutton
import time
import os
from PIL import Image, ImageTk

class MainInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("Menu Principal")
        self.root.geometry("650x180")

        # Centraliza os botões verticalmente na página
        self.root.columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        welcome_label = tk.Label(self.root, text="Bem-vindo ao bot de envio de mensagens automáticas da unoesc!",fg="blue", font=("Arial", 15, "bold"))
        welcome_label.pack(pady=20)

        self.button_contact = tk.Button(self.root, text="Mensagem para Contato", command=self.open_contact_interface, bg="#FFB700", fg="#5D1D88")
        self.button_contact.pack(pady=10)

        self.button_group = tk.Button(self.root, text="Mensagem para Grupo", command=self.open_group_interface, bg="#FFB700", fg="#5D1D88")
        self.button_group.pack(pady=10)        

    def open_contact_interface(self):
        self.root.withdraw()
        contact_interface = ContactInterface(self.root, self)

    def open_group_interface(self):
        self.root.withdraw()
        group_interface = GroupInterface(self.root, self)

class ContactInterface:
    def __init__(self, parent, main_interface):
        self.selected_file = None
        self.main_interface = main_interface  # Referência para a interface pai
        
        # Inicialização da janela principal do aplicativo
        self.parent = parent
        self.root = tk.Toplevel(self.parent)
        self.root.title("Bot.unoesc - Envio de Mensagens Automáticas")
        self.root.geometry("600x500")
        self.root.protocol("WM_DELETE_WINDOW", self.close_contact_interface) 

        # Criação do quadro dentro da janela
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        # Personalizar cores
        self.root.configure(bg="#f0f0f0")
        self.frame.configure(bg="#f0f0f0")

        # Rótulo para exibir a mensagem de aviso
        self.show_welcome_message()

        # Botão para voltar ao menu
        self.back_button = tk.Button(self.root, text="Voltar ao Menu", command=self.back_to_main_interface)
        self.back_button.pack()

        # Label para exibir o nome do arquivo
        self.file_label = tk.Label(self.frame, text="", bg="#e0e0e0", relief="sunken")
        self.file_label.place(x=226, y=7, width=110)
        
        # Botão para selecionar um arquivo
        self.select_button = tk.Button(self.frame, text="Selecionar Arquivo", command=self.select_file, bg="#007acc", fg="white")
        self.select_button.pack(pady=(30, 50))

        # Botão para selecionar uma imagem
        self.select_image_button = tk.Button(self.frame, text="Selecionar Imagem", command=self.select_image, bg="#007acc", fg="white")
        self.select_image_button.place(x=447, y=99)
        
        # Label para exibir o nome da imagem
        self.image_label = tk.Label(self.root, text="", relief="sunken", bg="#e0e0e0")
        self.image_label.place(x=470, y=95, width=110)
        
        # Rótulo para a caixa de texto
        self.message_label = tk.Label(self.frame, text="Digite sua mensagem aqui:", bg="#f0f0f0")
        self.message_label.pack()

        # Caixa de texto
        self.message_text = tk.Text(self.frame, height=10, width=70, bg="white", wrap="word")
        self.message_text.pack()
        self.message_text.insert("1.0","Use {código} para substituir os valores.\n\nExemplo:\n\nOlá {aluno}, estamos felizes em tê-lo no curso de {curso} {roseli}{roseli}{roseli}!")
          
        # Botão de legenda
        self.legend_button = tk.Button(self.frame, text="Legenda", command=self.show_legend, bg="#ffcc00")
        self.legend_button.place(x=0, y=99)

        # Botão para enviar as mensagens
        self.send_button = tk.Button(self.frame, text="Enviar Mensagens", command=self.validate_and_send_messages, bg="#4caf50", fg="white")
        self.send_button.pack(pady=(10,0))
        
        # Rótulo para exibir mensagens de aviso
        self.warning_label = tk.Label(self.frame, text="", fg="red")
        self.warning_label.pack()
        
        # Substitui emojis pelos placeholders
        self.emoji_mapping = {
            "{daniel}": "😊",
            "{arthur}": "😍",
            "{patriciane}": "😁",
            "{aline}": "😉",
            "{roseli}": "😄",
            "{vanessa}": "🤗",
            "{yuri}": "😜"
        }

        # Rótulo com a versão
        version_label = tk.Label(self.root, text="Versão 4.1", bg="#f0f0f0", fg="gray")
        version_label.pack(side="bottom", padx=10, pady=10, anchor="se")
        
        # Inicia a interface gráfica
        self.root.mainloop()

    def close_contact_interface(self):
        self.root.destroy()
        self.parent.deiconify()

    def show_welcome_message(self):
        welcome_message = (
            "Bot.unoesc - Envio de Mensagens Automáticas - contatos!\n"
            "Lembre-se de ajustar os cabeçalhos da planilha para que os dados sejam importados corretamente no programa:\n"
            "Nome do aluno, Nome do curso, Telefone ou Telefone 1."
        )
        messagebox.showinfo("Aviso de Boas-Vindas", welcome_message)

    def back_to_main_interface(self):
        self.root.destroy()
        self.main_interface.root.deiconify()

    def select_file(self):
        # Função para selecionar um arquivo .xlsx
        self.selected_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.selected_file:
            messagebox.showinfo("Arquivo Selecionado", "Arquivo selecionado com sucesso!")
            file_name = os.path.basename(self.selected_file)
            self.file_label.config(text=f"{file_name}")

    def select_image(self):
        # Função para selecionar uma imagem
        image_file = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png *.jpeg")])
        if image_file:
            self.image_path = image_file
            messagebox.showinfo("Imagem Selecionada", "Imagem selecionada com sucesso!")
            self.show_image()

            # Atualiza a label com o nome do arquivo
            file_name = os.path.basename(image_file)
            self.image_label.config(text=file_name)

    def show_image(self):
        if self.image_path:
            image = Image.open(self.image_path)
            image = image.resize((400, 400))

            # Cria uma nova janela para mostrar a imagem
            image_window = tk.Toplevel(self.root)
            image_window.title("Visualização de Imagem")

            # Converte a imagem para o formato do widget Label
            photo = ImageTk.PhotoImage(image)

            # Configura o widget Label para mostrar a imagem
            image_label = tk.Label(image_window, image=photo)
            image_label.photo = photo 
            image_label.pack()

            ok_button = tk.Button(image_window, text="OK", command=image_window.destroy, bg="#007acc", fg="white")
            ok_button.pack(pady=10)

    def validate_and_send_messages(self):
        # Função para validar e enviar as mensagens
        if self.selected_file is None:
            messagebox.showwarning("Aviso", "Selecione um arquivo antes de enviar as mensagens.")
            return

        try:
            df = pd.read_excel(self.selected_file)
        except pd.errors.EmptyDataError:
            messagebox.showerror("Erro", "O arquivo selecionado está vazio. Selecione um arquivo válido.")
            return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo selecionado: {e}")
            return

        custom_message = self.message_text.get("1.0", tk.END).strip()
        
        for emoji_placeholder, emoji_code in self.emoji_mapping.items():
            custom_message = custom_message.replace(emoji_placeholder, emoji_code)

        messages_to_send = []
        
        if 'Nome do aluno' in df.columns and ('Telefone' in df.columns or 'Telefone 1' in df.columns):
            for index, row in df.iterrows():
                full_name = row['Nome do aluno']
                aluno = ''.join([name.capitalize() for name in full_name.split()[:1]])
                
                if 'Nome do curso' in df.columns:
                    curso = row['Nome do curso']
                    curso = ' '.join([part.capitalize() for part in curso.split()])
                else:
                    curso = ""
                
                telefone = row.get('Telefone', row.get('Telefone 1', None))

                if telefone is None or pd.isnull(telefone):
                    messagebox.showwarning("Erro", f"Telefone não encontrado para {full_name}. Pulando mensagem.")
                    continue
                    
                telefone_numerico = re.sub(r'\D', '', str(telefone))
                
                if len(telefone_numerico) != 11:
                    messagebox.showwarning("Erro", f"Telefone de {full_name} com número incorreto de dígitos.")
                    continue
            
                telefone_formatado = '+55' + telefone_numerico

                message = custom_message.format(aluno=aluno, curso=curso)
                
                messages_to_send.append((full_name, telefone_formatado, message))

            self.show_validation_dialog(messages_to_send)
        else:
            messagebox.showwarning("Aviso", "As colunas 'aluno' e/ou 'telefone' não foram encontradas no arquivo.")

    def send_messages(self, messages_to_send, validation_dialog):
        # Função para enviar as mensagens
        validation_dialog.destroy()
        
        for aluno, telefone_formatado, message in messages_to_send:
            if hasattr(self, 'image_path'):
                kit.sendwhats_image(telefone_formatado, self.image_path, message, 20, 3)
            else:
                kit.sendwhatmsg_instantly(telefone_formatado, message, 15, 3)
        success_message = "Mensagens enviadas com sucesso!"
        messagebox.showinfo("Sucesso", success_message)
                 
    def show_validation_dialog(self, messages_to_send):
        # Mostra janela de validação
        validation_dialog = tk.Toplevel(self.root)
        validation_dialog.title("Validação de Mensagens")
        validation_dialog.attributes('-toolwindow', True)
        validation_textbox = tk.Text(validation_dialog, wrap="word")
        validation_textbox.pack(fill="both", expand=True)

        # Verifica as mensagens
        for aluno, telefone_formatado, message in messages_to_send:
            validation_textbox.insert("end", f"{aluno} ({telefone_formatado}): {message}\n{'-' * 80}\n")

        send_button = tk.Button(validation_dialog, text="Enviar Mensagens", command=lambda: self.send_messages(messages_to_send, validation_dialog), bg="#4caf50", fg="white")
        send_button.pack(pady=10)

        cancel_button = tk.Button(validation_dialog, text="Cancelar Envio", command=validation_dialog.destroy, bg="#f44336", fg="white")
        cancel_button.pack(pady=10)

    def show_legend(self):
        # Função para mostrar a legenda
        legend_window = tk.Toplevel(self.root)
        legend_window.title("Legenda")
        legend_window.geometry("400x300")

        legend_text = (
            "{aluno} = aparecerá o nome do aluno\n"
            "{curso} = aparecerá o nome do curso\n"
            "{daniel} = 😊 Emoji de sorriso\n"
            "{arthur} = 😍 Emoji de coração nos olhos\n"
            "{patriciane} = 😁 Emoji sorridente com olhos fechados\n"
            "{aline} = 😉 Emoji piscando\n"
            "{roseli} = 😄 Emoji feliz\n"
            "{vanessa} = 🤗 Emoji de abraço\n"
            "{yuri} = 😜 Emoji de língua para fora"
        )
        
        legend_textbox = tk.Text(legend_window, wrap="word")
        legend_textbox.pack(fill="both", expand=True)
        legend_textbox.insert("end", legend_text)

class GroupInterface:
    def __init__(self, parent, main_interface):
        self.selected_file = None
        self.main_interface = main_interface

        # Inicialização da janela principal do aplicativo
        self.parent = parent
        self.root = tk.Toplevel(self.parent)
        self.root.title("Bot.unoesc - Envio de Mensagens Automáticas")
        self.root.geometry("600x500")
        self.root.protocol("WM_DELETE_WINDOW", self.close_Group_interface)
        
        # Quadro dentro da janela
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        # Personalizar cores
        self.root.configure(bg="#f0f0f0")
        self.frame.configure(bg="#f0f0f0")

        # Rótulo de aviso para exibir a mensagem
        self.show_welcome_message()

        # Botão para voltar ao menu
        self.back_button = tk.Button(self.root, text="Voltar ao Menu", command=self.back_to_main_interface)
        self.back_button.pack()

        # Label para exibir o nome do arquivo
        self.file_label = tk.Label(self.frame, text="", bg="#e0e0e0", relief="sunken")
        self.file_label.place(x=226, y=7, width=110)
        
        # Botão para selecionar um arquivo
        self.select_button = tk.Button(self.frame, text="Selecionar Arquivo", command=self.select_file, bg="#007acc", fg="white")
        self.select_button.pack(pady=(30, 50))

        # Botão para selecionar uma imagem
        self.select_image_button = tk.Button(self.frame, text="Selecionar Imagem", command=self.select_image, bg="#007acc", fg="white")
        self.select_image_button.place(x=447, y=99)
        
        # Label para exibir o nome da imagem
        self.image_label = tk.Label(self.root, text="", relief="sunken", bg="#e0e0e0")
        self.image_label.place(x=470, y=95, width=110)
        
        # Rótulo para a caixa de texto
        self.message_label = tk.Label(self.frame, text="Digite sua mensagem aqui:", bg="#f0f0f0")
        self.message_label.pack()

        # Caixa de texto
        self.message_text = tk.Text(self.frame, height=10, width=70, bg="white", wrap="word")
        self.message_text.pack()
        self.message_text.insert("1.0","Use {código} para substituir os valores.\n\nExemplo:\n\nOlá alunos hoje começará as aulas de {curso}, cuidado com o bot do café.{roseli}{roseli}{roseli}!")
          
        # Botão de legenda
        self.legend_button = tk.Button(self.frame, text="Legenda", command=self.show_legend, bg="#ffcc00")
        self.legend_button.place(x=0, y=99)

        # Botão para enviar as mensagens
        self.send_button = tk.Button(self.frame, text="Enviar Mensagens", command=self.validate_and_send_messages, bg="#4caf50", fg="white")
        self.send_button.pack(pady=(10,0))
        
        # Rótulo para exibir mensagens de aviso
        self.warning_label = tk.Label(self.frame, text="", fg="red")
        self.warning_label.pack()
        
        # Substitui emojis pelos placeholders
        self.emoji_mapping = {
            "{daniel}": "😊",
            "{arthur}": "😍",
            "{patriciane}": "😁",
            "{aline}": "😉",
            "{roseli}": "😄",
            "{vanessa}": "🤗",
            "{yuri}": "😜"
        }

        # Rótulo com a versão
        version_label = tk.Label(self.root, text="Versão 4.1", bg="#f0f0f0", fg="gray")
        version_label.pack(side="bottom", padx=10, pady=10, anchor="se")
        
        # Inicia a interface gráfica
        self.root.mainloop()

    def close_Group_interface(self):
        self.root.destroy()
        self.parent.deiconify()

    def show_welcome_message(self):
        welcome_message = (
            "Bot.unoesc - Envio de Mensagens Automáticas - grupo!\n"
            "Lembre-se de ajustar os cabeçalhos da planilha para que os dados sejam importados corretamente no programa:\n"
            "'Nome do grupo'\n'Nome do curso'\n'Horário'."
        )
        messagebox.showinfo("Aviso de Boas-Vindas", welcome_message)

    def back_to_main_interface(self):
        self.root.destroy() 
        self.main_interface.root.deiconify()

    def select_file(self):
        # Função para selecionar um arquivo Excel (.xlsx)
        self.selected_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.selected_file:
            messagebox.showinfo("Arquivo Selecionado", "Arquivo selecionado com sucesso!")
            file_name = os.path.basename(self.selected_file)
            self.file_label.config(text=f"{file_name}")
            
    def select_image(self):
        # Função para selecionar uma imagem
        image_file = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png *.jpeg")])
        if image_file:
            self.image_path = image_file
            messagebox.showinfo("Imagem Selecionada", "Imagem selecionada com sucesso!")
            self.show_image()

            # Atualizar a label fixa com o nome do arquivo
            file_name = os.path.basename(image_file)
            self.image_label.config(text=file_name)

    def show_image(self):
        if self.image_path:
            image = Image.open(self.image_path)
            image = image.resize((400, 400))  

            # Janela para mostrar a imagem
            image_window = tk.Toplevel(self.root)
            image_window.title("Visualização de Imagem")

            # Converte a imagem para o formato exibível pelo widget Label
            photo = ImageTk.PhotoImage(image)

            # Widget Label para mostrar a imagem
            image_label = tk.Label(image_window, image=photo)
            image_label.photo = photo  # Mantenha uma referência para evitar que a imagem seja liberada da memória
            image_label.pack()

            ok_button = tk.Button(image_window, text="OK", command=image_window.destroy, bg="#007acc", fg="white")
            ok_button.pack(pady=10)

    def validate_and_send_messages(self):
        # Função para validar e enviar as mensagens
        if self.selected_file is None:
            messagebox.showwarning("Aviso", "Selecione um arquivo antes de enviar as mensagens.")
            return

        try:
            df = pd.read_excel(self.selected_file)
        except pd.errors.EmptyDataError:
            messagebox.showerror("Erro", "O arquivo selecionado está vazio. Selecione um arquivo válido.")
            return
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo selecionado: {e}")
            return

        custom_message = self.message_text.get("1.0", tk.END).strip()
        
        for emoji_placeholder, emoji_code in self.emoji_mapping.items():
            custom_message = custom_message.replace(emoji_placeholder, emoji_code)

        messages_to_send = []
        
        if 'Nome do grupo' in df.columns:
            for index, row in df.iterrows():
                grupo = row['Nome do grupo']
                
                if 'Horário' in df.columns:
                    horario = row['Horário']
                else:
                    horario = "" 
                    
                if 'Nome do curso' in df.columns:
                    curso = row['Nome do curso']
                else:
                    curso = ""
                
           
                message = custom_message.format(curso=curso, horario=horario)
                
                messages_to_send.append((grupo, curso, message))
                print(grupo)

            self.show_validation_dialog(messages_to_send)
            print("\n", grupo)
        else:
            messagebox.showwarning("Aviso", "A coluna 'grupo' não foi encontrada no arquivo.")

    def send_messages(self, messages_to_send, validation_dialog):
        # Função para enviar as mensagens
        validation_dialog.destroy()
        
        for grupo, curso, message in messages_to_send:
            if hasattr(self, 'image_path'):
                print(message)
                kit.sendwhats_image(grupo, self.image_path, message, 30, 3)
            else:
                print(message)
                kit.sendwhatmsg_to_group_instantly(grupo, message, 20, 3)
        success_message = "Mensagens enviadas com sucesso!"
        messagebox.showinfo("Sucesso", success_message)
                 
    def show_validation_dialog(self, messages_to_send):
        # Mostra janela de validação de mensagens
        validation_dialog = tk.Toplevel(self.root)
        validation_dialog.title("Validação de Mensagens")
        validation_dialog.attributes('-toolwindow', True)
        validation_textbox = tk.Text(validation_dialog, wrap="word")
        validation_textbox.pack(fill="both", expand=True)

        # Verifica as mensagens para a caixa de texto
        for grupo, curso, message in messages_to_send:
            validation_textbox.insert("end", f"{curso} ({grupo}): {message}\n{'-' * 80}\n")

        send_button = tk.Button(validation_dialog, text="Enviar Mensagens", command=lambda: self.send_messages(messages_to_send, validation_dialog), bg="#4caf50", fg="white")
        send_button.pack(pady=10)

        cancel_button = tk.Button(validation_dialog, text="Cancelar Envio", command=validation_dialog.destroy, bg="#f44336", fg="white")
        cancel_button.pack(pady=10)

    def show_legend(self):
        # Função para mostrar a legenda em uma janela separada
        legend_window = tk.Toplevel(self.root)
        legend_window.title("Legenda")
        legend_window.geometry("400x300")

        legend_text = (
            "{grupo} = aparecerá o nome do grupo\n"
            "{componente} = aparecerá o nome do componente\n"
            "{daniel} = 😊 Emoji de sorriso\n"
            "{arthur} = 😍 Emoji de coração nos olhos\n"
            "{patriciane} = 😁 Emoji sorridente com olhos fechados\n"
            "{aline} = 😉 Emoji piscando\n"
            "{roseli} = 😄 Emoji feliz\n"
            "{vanessa} = 🤗 Emoji de abraço\n"
            "{yuri} = 😜 Emoji de língua para fora"
        )
        
        legend_textbox = tk.Text(legend_window, wrap="word")
        legend_textbox.pack(fill="both", expand=True)
        legend_textbox.insert("end", legend_text)

def main():
    root = tk.Tk()
    app = MainInterface(root)
    root.mainloop()

if __name__ == "__main__":
    main()