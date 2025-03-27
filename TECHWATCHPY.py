import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from tkinter.simpledialog import askstring
import pandas as pd
from datetime import datetime
import logging
import json
import os
from tkinter.ttk import Progressbar

# Configuração do Logging
logging.basicConfig(filename='inventory.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Carregar configurações
CONFIG_FILE = "config.json"
DEFAULT_CONFIG = {"geometry": "1200x800", "last_dir": "."}

def load_config():
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return DEFAULT_CONFIG
    except json.JSONDecodeError:
        logging.error("Erro ao decodificar o arquivo de configuração. Usando configurações padrão.")
        return DEFAULT_CONFIG

def save_config(config):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f)
    except Exception as e:
        logging.error(f"Erro ao salvar configurações: {e}")

config = load_config()

class ItemDialog(tk.Toplevel):
    # ... (Seu código ItemDialog permanece praticamente o mesmo)
    def __init__(self, parent, title, fields, item_data=None):
        super().__init__(parent)
        self.title(title)
        self.fields = fields
        self.entries = {}
        self.item_data = item_data  # Dados do item para edição
        self.create_widgets()
        self.transient(parent)
        self.grab_set()
        self.focus_set()

    def create_widgets(self):
        frame = ttk.LabelFrame(self, text=self.title, padding=10)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        for label_text in self.fields:
            tk.Label(frame, text=label_text).pack(pady=2, anchor=tk.W)
            entry = tk.Entry(frame)
            entry.pack(pady=2, fill=tk.X)
            self.entries[label_text] = entry
            if self.item_data:
                entry.insert(0, self.item_data.get(label_text, ''))

        button_frame = tk.Frame(self)
        button_frame.pack(pady=10)
        add_button = tk.Button(button_frame, text="Salvar" if self.item_data else "Adicionar", command=self.on_add)
        add_button.pack(side=tk.LEFT, padx=5)

        cancel_button = tk.Button(button_frame, text="Cancelar", command=self.destroy)
        cancel_button.pack(side=tk.LEFT, padx=5)

    def validate_date(self, date_str):
        try:
            datetime.strptime(date_str, "%d/%m/%Y")
            return True
        except ValueError:
            return False

    def is_number(self, value):
        try:
            float(value)
            return True
        except ValueError:
            return False

    def on_add(self):
        details = {}
        for label, entry in self.entries.items():
            value = entry.get().strip()
            if label in ("Número de Patrimônio", "Nome", "Setor"):
                if not value:
                    messagebox.showerror("Erro", f"O campo '{label}' é obrigatório.")
                    return
            if label == "Data de Compra (DD/MM/YYYY)" and value:
                if not self.validate_date(value):
                    messagebox.showerror("Erro", "Formato de data inválido (DD/MM/YYYY).")
                    return
                try:
                    purchase_date = datetime.strptime(value, "%d/%m/%Y")
                    if purchase_date > datetime.now():
                        messagebox.showerror("Erro", "A data de compra não pode ser futura.")
                        return
                except ValueError:
                    messagebox.showerror("Erro", "Data inválida.")
                    return
            if label == "Memória RAM" and value:
                if not self.is_number(value):
                    messagebox.showerror("Erro", "Memória RAM deve ser um número.")
                    return

            details[label] = value

        self.add_item(details)
        self.destroy()
        messagebox.showinfo("Sucesso", f"{self.item_type} {'editado' if self.item_data else 'adicionado'} com sucesso!")

    def add_item(self, details):
        pass

class DeviceDialog(ItemDialog):
    # ... (Seu código DeviceDialog permanece praticamente o mesmo)
    def __init__(self, parent, device_type, item_data=None):
        self.device_type = device_type
        fields = [
            "Número de Patrimônio", "Nome", "Modelo", "Setor", "Usuário",
            "Memória RAM", "S.O.", "Processador", "Data de Compra (DD/MM/YYYY)",
            "Última Manutenção Feita", "Observações"
        ]
        super().__init__(parent, f"Detalhes do {device_type}", fields, item_data)
        self.item_type = device_type

    def add_item(self, details):
        details["Tipo"] = self.device_type
        details["Status"] = "Ativo"
        if self.item_data:
            self.master.app.update_item_in_inventory(self.item_data["Número de Patrimônio"], details)
        else:
            self.master.app.add_item_to_inventory(details)

class NotebookDialog(DeviceDialog):
    def __init__(self, parent, item_data=None):
        super().__init__(parent, "Notebook", item_data)

class DesktopDialog(DeviceDialog):
    def __init__(self, parent, item_data=None):
        super().__init__(parent, "Desktop", item_data)

class MonitorDialog(ItemDialog):
    # ... (Seu código MonitorDialog permanece praticamente o mesmo)
    def __init__(self, parent, item_data=None):
        fields = [
            "Número de Patrimônio", "Nome", "Modelo", "Setor", "Usuário",
            "Tamanho da Tela", "Tipo de Conexão", "Data de Compra (DD/MM/YYYY)", "Observações"
        ]
        super().__init__(parent, "Detalhes do Monitor", fields, item_data)
        self.item_type = "Monitor"

    def add_item(self, details):
        details["Tipo"] = "Monitor"
        details["Status"] = "Ativo"
        if self.item_data:
            self.master.app.update_item_in_inventory(self.item_data["Número de Patrimônio"], details)
        else:
            self.master.app.add_item_to_inventory(details)

class OtherItemDialog(ItemDialog):
    # ... (Seu código OtherItemDialog permanece praticamente o mesmo)
    def __init__(self, parent, item_data=None):
        fields = [
            "Número de Patrimônio", "Nome", "Modelo", "Setor", "Usuário",
            "Data de Compra (DD/MM/YYYY)", "Observações"
        ]
        super().__init__(parent, "Detalhes de Outro Item", fields, item_data)
        self.item_type = "Outros"

    def add_item(self, details):
        details["Tipo"] = "Outros"
        details["Status"] = "Ativo"
        if self.item_data:
            self.master.app.update_item_in_inventory(self.item_data["Número de Patrimônio"], details)
        else:
            self.master.app.add_item_to_inventory(details)

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventário de TI")

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))

        self.root.geometry(config["geometry"])
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.data = pd.DataFrame(columns=[
            "Número de Patrimônio", "Tipo", "Nome", "Modelo", "Setor", "Usuário",
            "Memória RAM", "S.O.", "Processador", "Data de Compra",
            "Última Manutenção Feita", "Observações", "Status"
        ])
        self.create_widgets()
        self.sort_order = {}
        self.current_file = None # Adicionado para rastrear o arquivo aberto

        self.root.app = self
        logging.info("Aplicativo iniciado.")

    def create_widgets(self):
        menubar = tk.Menu(self.root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Salvar", command=self.save_to_excel, accelerator="Ctrl+S")
        filemenu.add_command(label="Carregar", command=self.load_from_excel)
        filemenu.add_command(label="Exportar CSV", command=self.export_to_csv)
        filemenu.add_separator()
        filemenu.add_command(label="Sair", command=self.on_closing)
        menubar.add_cascade(label="Arquivo", menu=filemenu)
        self.root.config(menu=menubar)

        self.root.bind("<Control-s>", self.save_to_excel)

        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        self.add_button = tk.Button(button_frame, text="Adicionar Item", command=self.show_item_type_dialog)
        self.add_button.pack(side=tk.LEFT, padx=5)
        self.root.bind("<Control-n>", lambda event: self.show_item_type_dialog())

        self.save_button = tk.Button(button_frame, text="Salvar Tabela", command=self.save_to_excel)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.load_button = tk.Button(button_frame, text="Carregar Tabela", command=self.load_from_excel)
        self.load_button.pack(side=tk.LEFT, padx=5)

        filter_frame = tk.Frame(button_frame)
        filter_frame.pack(side=tk.LEFT, padx=5)

        self.filter_label = tk.Label(filter_frame, text="Filtrar:")
        self.filter_label.pack(side=tk.LEFT, padx=2)

        self.filter_entry = tk.Entry(filter_frame)
        self.filter_entry.pack(side=tk.LEFT, padx=2)

        self.filter_button = tk.Button(filter_frame, text="Filtrar", command=self.filter_items)
        self.filter_button.pack(side=tk.LEFT, padx=2)

        self.clear_filter_button = tk.Button(filter_frame, text="Limpar Filtro", command=self.clear_filter)
        self.clear_filter_button.pack(side=tk.LEFT, padx=2)

        self.delete_button = tk.Button(button_frame, text="Excluir Item", command=self.delete_item)
        self.delete_button.pack(side=tk.LEFT, padx=5)

        self.edit_button = tk.Button(button_frame, text="Editar Item", command=self.edit_item)
        self.edit_button.pack(side=tk.LEFT, padx=5)

        frame = tk.Frame(self.root)
        frame.pack(pady=5, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(frame, columns=list(self.data.columns), show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        for column in self.tree["columns"]:
            self.tree.heading(column, text=column, command=lambda col=column: self.sort_column(col))
            self.tree.column(column, width=150, stretch=tk.YES)  # Configuração inicial da largura

        self.status_bar = tk.Label(self.root, text="Pronto", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.progress_bar = Progressbar(self.root, orient=tk.HORIZONTAL, mode='indeterminate')

    def show_item_type_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Selecione o Tipo de Item")
        dialog.geometry("250x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_set()
        def on_select(item_type):
            dialog.destroy()
            if item_type == "Notebook":
                NotebookDialog(self.root)
            elif item_type == "Desktop":
                DesktopDialog(self.root)
            elif item_type == "Monitor":
                MonitorDialog(self.root)
            elif item_type == "Outros":
                OtherItemDialog(self.root)
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Notebook", command=lambda: on_select("Notebook")).pack(pady=5,fill = tk.X)
        ttk.Button(button_frame, text="Desktop", command=lambda: on_select("Desktop")).pack(pady=5,fill = tk.X)
        ttk.Button(button_frame, text="Monitor", command=lambda: on_select("Monitor")).pack(pady=5,fill = tk.X)
        ttk.Button(button_frame, text="Outros", command=lambda: on_select("Outros")).pack(pady=5,fill = tk.X)
        dialog.wait_window()

    def add_item_to_inventory(self, details):
        if self.data["Número de Patrimônio"].isin([details["Número de Patrimônio"]]).any():
            messagebox.showerror("Erro", "Número de Patrimônio já existe. Insira um valor único.")
            return

        self.data = pd.concat([self.data, pd.DataFrame([details])], ignore_index=True)
        self.update_treeview()
        logging.info(f"Item adicionado: {details['Tipo']} - {details['Nome']}")
        self.update_status("Item adicionado com sucesso!")

    def update_item_in_inventory(self, asset_number, details):
        index = self.data[self.data["Número de Patrimônio"] == asset_number].index[0]
        if not index.empty:
            self.data.loc[index[0]] = list(details.values())
            self.update_treeview()
            logging.info(f"Item atualizado: {details['Tipo']} - {details['Nome']}")
            self.update_status("Item atualizado com sucesso!")
        else:
            messagebox.showerror("Erro", "Número de Patrimônio não encontrado.")
            logging.error(f"Erro ao atualizar item: {asset_number} não encontrado.")
            self.update_status("Erro ao atualizar item.")

    def edit_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_data = self.tree.item(selected_item)["values"]
            asset_number = item_data[0]
            item = self.data[self.data["Número de Patrimônio"] == asset_number].iloc[0].to_dict()
            item_type = item["Tipo"]

            if item_type == "Notebook":
                NotebookDialog(self.root, item)
            elif item_type == "Desktop":
                DesktopDialog(self.root, item)
            elif item_type == "Monitor":
                MonitorDialog(self.root, item)
            elif item_type == "Outros":
                OtherItemDialog(self.root, item)
        else:
            messagebox.showinfo("Aviso", "Selecione um item para editar.")

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for row in self.data.to_numpy().tolist():
            self.tree.insert("", "end", values=row)
        self.update_status("Tabela atualizada.")
        self.adjust_column_widths()

    def delete_item(self):
        selected_item = self.tree.selection()
        if selected_item:
            confirm = messagebox.askyesno("Confirmação", "Deseja realmente excluir o item?")
            if confirm:
                item_data = self.tree.item(selected_item)["values"]
                asset_number = item_data[0]
                self.data = self.data[self.data["Número de Patrimônio"] != asset_number]
                self.tree.delete(selected_item)
                self.update_treeview()
                logging.info(f"Item excluído: {asset_number}")
                messagebox.showinfo("Sucesso", "Item excluído com sucesso!")
                self.update_status("Item excluído.")
        else:
            messagebox.showinfo("Aviso", "Selecione um item para excluir.")

    def filter_items(self):
        filter_value = self.filter_entry.get().strip().lower()
        if filter_value:
            filtered_data = self.data[self.data.apply(lambda row: any(filter_value in str(item).lower() for item in row), axis=1)]
            self.update_treeview_with_data(filtered_data)
            self.update_status(f"Filtrando por: '{filter_value}'")
        else:
            self.update_treeview()
            self.update_status("Filtro removido.")

    def clear_filter(self):
        self.filter_entry.delete(0, tk.END)
        self.update_treeview()
        self.update_status("Filtro removido.")

    def update_treeview_with_data(self, data):
        self.tree.delete(*self.tree.get_children())
        for row in data.to_numpy().tolist():
            self.tree.insert("", "end", values=row)
        self.adjust_column_widths()

    def save_to_excel(self, event=None):
        config = load_config()
        initialdir = config.get("last_dir", ".")
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx")],
                                                initialdir=initialdir)
        if file_path:
            try:
                self.start_progress()
                self.data.to_excel(file_path, index=False)
                logging.info(f"Tabela salva em: {file_path}")
                self.update_status(f"Tabela salva em: {file_path}")
                messagebox.showinfo("Sucesso", f"Tabela salva em: {file_path}")
                self.current_file = file_path # Atualiza o arquivo atual
                config["last_dir"] = os.path.dirname(file_path)
                save_config(config)
            except Exception as e:
                logging.error(f"Erro ao salvar o arquivo: {e}")
                messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")
                self.update_status("Erro ao salvar o arquivo.")
            finally:
                self.stop_progress()

    def load_from_excel(self):
        config = load_config()
        initialdir = config.get("last_dir", ".")
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")],
                                                initialdir=initialdir)
        if file_path:
            try:
                self.start_progress()
                self.data = pd.read_excel(file_path)
                self.update_treeview()
                logging.info(f"Tabela carregada de: {file_path}")
                self.update_status(f"Tabela carregada de: {file_path}")
                messagebox.showinfo("Sucesso", "Tabela carregada com sucesso!")
                self.current_file = file_path # Atualiza o arquivo atual
                config["last_dir"] = os.path.dirname(file_path)
                save_config(config)
            except FileNotFoundError:
                logging.error("Arquivo não encontrado.")
                messagebox.showerror("Erro", "Arquivo não encontrado.")
                self.update_status("Arquivo não encontrado.")
            except Exception as e:
                logging.error(f"Erro ao carregar o arquivo: {e}")
                messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")
                self.update_status("Erro ao carregar o arquivo.")
            finally:
                self.stop_progress()

    def export_to_csv(self):
        config = load_config()
        initialdir = config.get("last_dir", ".")
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                filetypes=[("CSV files", "*.csv")],
                                                initialdir=initialdir)
        if file_path:
            try:
                self.start_progress()
                self.data.to_csv(file_path, index=False)
                logging.info(f"Tabela exportada para CSV em: {file_path}")
                self.update_status(f"Tabela exportada para CSV em: {file_path}")
                messagebox.showinfo("Sucesso", f"Tabela exportada para CSV em: {file_path}")
                config["last_dir"] = os.path.dirname(file_path)
                save_config(config)
            except Exception as e:
                logging.error(f"Erro ao exportar para CSV: {e}")
                messagebox.showerror("Erro", f"Erro ao exportar para CSV: {e}")
                self.update_status("Erro ao exportar para CSV.")
            finally:
                self.stop_progress()

    def sort_column(self, column):
        if column not in self.data.columns:
            messagebox.showwarning("Aviso", f"A coluna '{column}' não existe no DataFrame.")
            return

        if column in self.sort_order:
            self.sort_order[column] = not self.sort_order[column]
        else:
            self.sort_order[column] = True

        ascending = self.sort_order[column]

        try:
            self.data = self.data.sort_values(by=[column], ascending=ascending, na_position='last')
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ordenar a coluna '{column}': {e}")
            return

        self.update_treeview()
        sort_direction = "ascendente" if ascending else "descendente"
        self.update_status(f"Ordenado por '{column}' ({sort_direction}).")

    def update_status(self, message):
        self.status_bar.config(text=message)
        logging.info(message)

    def on_closing(self):
        config["geometry"] = self.root.winfo_geometry()
        root.geometry()
        save_config(config)

        if messagebox.askokcancel("Sair", "Deseja salvar as alterações antes de sair?"):
            self.save_to_excel()
        self.root.destroy()
        logging.info("Aplicativo encerrado.")

    def adjust_column_widths(self):
        for col in self.tree["columns"]:
            max_width = 0
            for item in self.tree.get_children():
                if self.tree.set(item, col):
                    max_width = max(max_width, tk.font.Font().measure(self.tree.set(item, col)))
            self.tree.column(col, width=max_width + 20)  # Adiciona algum padding

    def start_progress(self):
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)
        self.progress_bar.start()

    def stop_progress(self):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()