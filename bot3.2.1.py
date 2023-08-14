# Esta é a versão 3.2.1 do Bot.unoesc - Envio de Mensagens Automáticas

import pandas as pd
import re
import pywhatkit as kit
import tkinter as tk
import multiprocessing
from tkinter import filedialog
from tkinter import messagebox

class StopButtonApp:
    def __init__(self, app_process):
        self.app_process = app_process
        self.root = tk.Tk()
        self.root.title("Parar Envio de Mensagens")
        self.root.geometry("300x100")

        self.status_label = tk.Label(self.root, text="Clique no botão para parar o envio de mensagens", fg="blue")
        self.status_label.pack(pady=20)

        self.stop_button = tk.Button(self.root, text="Parar Envio", command=self.stop_sending, bg="#f44336", fg="white")
        self.stop_button.pack()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()

    def stop_sending(self):
        self.app_process.terminate()
        self.root.destroy()

    def on_closing(self):
        self.stop_sending()

class AutoMessageSenderApp:
    def __init__(self):
        self.selected_file = None

        # Inicialização da janela principal do aplicativo
        self.root = tk.Tk()
        self.root.title("Bot.unoesc - Envio de Mensagens Automáticas")
        self.root.geometry("600x400")  # Definir o tamanho da janela

        # Criação do quadro dentro da janela
        self.frame = tk.Frame(self.root)
        self.frame.pack(padx=20, pady=20)

        # Personalizar cores
        self.root.configure(bg="#f0f0f0")  # Cor de fundo da janela
        self.frame.configure(bg="#f0f0f0")  # Cor de fundo do quadro

        # Botão para selecionar um arquivo
        self.select_button = tk.Button(self.frame, text="Selecionar Arquivo", command=self.select_file, bg="#007acc", fg="white")
        self.select_button.pack(pady=10)  # Espaçamento entre o botão e outros widgets

        # variável para o status
        self.status = []
        
        # Rótulo para a caixa de texto da mensagem
        self.message_label = tk.Label(self.frame, text="Digite sua mensagem aqui:", bg="#f0f0f0")
        self.message_label.pack()

        # Caixa de texto para a mensagem
        self.message_text = tk.Text(self.frame, height=5, width=50, bg="white")  # Cor de fundo da caixa de texto
        self.message_text.pack()
        self.message_text.insert("1.0", "Olá {aluno}, estamos felizes em tê-lo no curso de {curso}!")

        # Rótulo de instruções para o usuário
        self.instructions_label = tk.Label(self.frame, text="Use {aluno} e {curso} nos códigos de formatação para substituir os valores.", bg="#f0f0f0")
        self.instructions_label.pack()

        # Botão para enviar as mensagens
        self.send_button = tk.Button(self.frame, text="Enviar Mensagens", command=self.validate_and_send_messages, bg="#4caf50", fg="white")
        self.send_button.pack(pady=10)
        
        # Botão para exibir o status
        self.show_status_button = tk.Button(self.frame, text="Exibir Status", command=self.show_status, bg="#ffa500", fg="white")
        self.show_status_button.pack(pady=10)

        # Adicionar rótulo com a versão no canto inferior direito
        version_label = tk.Label(self.root, text="Versão 3.2.1", bg="#f0f0f0", fg="gray")
        version_label.pack(side="bottom", padx=10, pady=10, anchor="se")  # Posicionar no canto inferior direito
        
        # Iniciar a interface gráfica
        self.root.mainloop()

    def select_file(self):
        # Função para selecionar um arquivo Excel (.xlsx)
        self.selected_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.selected_file:
            # Exibir mensagem de sucesso se o arquivo for selecionado
            messagebox.showinfo("Arquivo Selecionado", "Arquivo selecionado com sucesso!")

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

        messages_to_send = []

        if 'Nome do aluno' in df.columns and 'Nome do curso' in df.columns and ('Telefone' in df.columns or 'Telefone 1' in df.columns):
            for index, row in df.iterrows():
                full_name = row['Nome do aluno']
                aluno = ''.join([name.capitalize() for name in full_name.split()[:1]])
                curso = row['Nome do curso']
                curso = ' '.join([part.capitalize() for part in curso.split()])
                
                if 'Telefone' in df.columns:
                    telefone = row['Telefone']
                elif 'Telefone 1' in df.columns:
                    telefone = row['Telefone 1']

                telefone_numerico = re.sub(r'\D', '', telefone)
                telefone_formatado = '+55' + telefone_numerico

                message = custom_message.format(aluno=aluno, curso=curso)

                messages_to_send.append((aluno, telefone_formatado, message))

            self.show_validation_dialog(messages_to_send)
        else:
            messagebox.showwarning("Aviso", "As colunas 'aluno', 'curso' e/ou 'telefone' não foram encontradas no arquivo.")

    def send_messages(self, messages_to_send, validation_dialog):
        # Função para enviar as mensagens
        validation_dialog.destroy()  # Fechar a janela de validação

        for aluno, telefone_formatado, message in messages_to_send:
            try:
                kit.sendwhatmsg_instantly(telefone_formatado, message)
                self.status.append((aluno, telefone_formatado, True))  # Armazenar o status de sucesso
            except Exception as send_exception:
                self.status.append((aluno, telefone_formatado, False))  # Armazenar o status de erro

        success_message = "Mensagens enviadas com sucesso!"
        messagebox.showinfo("Sucesso", success_message)
            
    def show_status(self):
        # Função para exibir o status das mensagens enviadas
        status_window = tk.Toplevel(self.root)
        status_window.title("Status de Envio")
        status_window.geometry("400x300")

        status_text = tk.Text(status_window, height=15, width=50)
        status_text.pack(padx=20, pady=20)

        for aluno, telefone, enviado in self.status:
            status_text.insert(tk.END, f"Aluno: {aluno}\nTelefone: {telefone}\nEnviado: {'Sim' if enviado else 'Não'}\n\n")

        status_text.config(state=tk.DISABLED)  # Torna o texto somente leitura
        
    def show_validation_dialog(self, messages_to_send):
        # Mostrar janela de validação de mensagens
        validation_dialog = tk.Toplevel(self.root)
        validation_dialog.title("Validação de Mensagens")
        
        # Maximizar a janela de validação
        validation_dialog.attributes('-toolwindow', True)

        validation_textbox = tk.Text(validation_dialog, wrap="word")
        validation_textbox.pack(fill="both", expand=True)

        # Adicionar mensagens para a caixa de texto
        for aluno, telefone_formatado, message in messages_to_send:
            validation_textbox.insert("end", f"{aluno} ({telefone_formatado}): {message}\n{'-' * 80}\n")

        send_button = tk.Button(validation_dialog, text="Enviar Mensagens", command=lambda: self.send_messages(messages_to_send, validation_dialog), bg="#4caf50", fg="white")
        send_button.pack(pady=10)

        cancel_button = tk.Button(validation_dialog, text="Cancelar Envio", command=validation_dialog.destroy, bg="#f44336", fg="white")
        cancel_button.pack(pady=10)

if __name__ == "__main__":
    # Inicializar a instância do aplicativo em um processo
    app_process = multiprocessing.Process(target=AutoMessageSenderApp)
    app_process.start()

    # Inicializar a instância do botão parar em outro processo
    stop_app = StopButtonApp(app_process)