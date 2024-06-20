import tkinter as tk
from tkinter import filedialog, messagebox
from tempfile import NamedTemporaryFile
from shutil import move
from odf.opendocument import load, OpenDocumentText
from odf.text import P, Span

class AplicacaoGUI:
    def __init__(self, root):
        self.root = root
        self.texto_original = tk.Text(self.root, height=10, width=50)
        self.entrada_cabecalho = tk.Entry(self.root, width=50)
        self.entrada_rodape = tk.Entry(self.root, width=50)

        self.criar_interface()

    def abrir_arquivo(self):
        filepath = filedialog.askopenfilename(filetypes=[("Arquivos ODF", "*.odt")])
        if filepath:
            doc = load(filepath)
            conteudo_original = self.get_text_from_odt(doc)
            self.texto_original.delete('1.0', tk.END)
            self.texto_original.insert(tk.END, conteudo_original)

    def get_text_from_odt(self, doc):
        text_elements = []
        for content in doc.getElementsByType(P):
            for node in content.childNodes:
                if isinstance(node, Span):
                    text_elements.append(node.text)
        return '\n'.join(text_elements)

    def processar_arquivo(self):
        novo_cabecalho = self.entrada_cabecalho.get()
        novo_rodape = self.entrada_rodape.get()

        conteudo_original = self.texto_original.get('1.0', tk.END)

        conteudo_modificado = f"{novo_cabecalho}\n\n{conteudo_original}\n\n{novo_rodape}"

        filepath = filedialog.asksaveasfilename(defaultextension=".odt",
                                                filetypes=[("Arquivos ODF", "*.odt")])
        if filepath:
            temp_file = NamedTemporaryFile(delete=False, suffix=".odt")
            temp_filename = temp_file.name
            temp_file.close()

            doc = load(filepath)
            self.update_odt_content(doc, conteudo_modificado)

            doc.save(temp_filename)
            move(temp_filename, filepath)

            messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")

    def update_odt_content(self, doc, new_content):
        # Limpar o conteúdo atual do documento ODT
        for element in doc.getElementsByType(P):
            for node in element.childNodes:
                element.removeChild(node)

        # Inserir novo conteúdo
        new_paragraphs = new_content.split('\n')
        for para_text in new_paragraphs:
            paragraph = P()
            span = Span(text=para_text)
            paragraph.appendChild(span)
            doc.text.addElement(paragraph)

    def criar_interface(self):
        self.root.title("Manipulação de Arquivo")

        btn_abrir = tk.Button(self.root, text="Abrir Arquivo", command=self.abrir_arquivo)
        btn_abrir.pack(pady=10)

        self.texto_original.pack(pady=10)

        lbl_cabecalho = tk.Label(self.root, text="Novo Cabeçalho:")
        lbl_cabecalho.pack()

        self.entrada_cabecalho.pack()

        lbl_rodape = tk.Label(self.root, text="Novo Rodapé:")
        lbl_rodape.pack()

        self.entrada_rodape.pack()

        btn_processar = tk.Button(self.root, text="Processar e Salvar", command=self.processar_arquivo)
        btn_processar.pack(pady=10)

def main():
    root = tk.Tk()
    app = AplicacaoGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

