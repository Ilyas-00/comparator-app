import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class ExcelComparatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparaison de spécifications NVRAM")

        self.frame_choose_files = tk.Frame(self.root, padx=10, pady=10)
        self.frame_choose_files.pack()

        self.frame_choose_columns = tk.Frame(self.root, padx=10, pady=10)
        self.frame_choose_columns.pack()

        self.frame_output = tk.Frame(self.root, padx=10, pady=10)
        self.frame_output.pack(expand=True, fill=tk.BOTH)

        self.frame_buttons = tk.Frame(self.root, padx=10, pady=10)
        self.frame_buttons.pack()

        self.label_spec1 = tk.Label(self.frame_choose_files, text="Fichier Excel 1:")
        self.label_spec1.grid(row=0, column=0, padx=5, pady=5)

        self.label_spec2 = tk.Label(self.frame_choose_files, text="Fichier Excel 2:")
        self.label_spec2.grid(row=1, column=0, padx=5, pady=5)

        self.label_column1 = tk.Label(self.frame_choose_columns, text="Colonne à comparer (Fichier 1):")
        self.label_column1.grid(row=0, column=0, padx=5, pady=5)

        self.label_column2 = tk.Label(self.frame_choose_columns, text="Colonne à comparer (Fichier 2):")
        self.label_column2.grid(row=1, column=0, padx=5, pady=5)

        self.column1_var = tk.StringVar()
        self.column2_var = tk.StringVar()

        self.entry_spec1 = ttk.Entry(self.frame_choose_files, width=30)
        self.entry_spec1.grid(row=0, column=1, padx=5, pady=5)

        self.entry_spec2 = ttk.Entry(self.frame_choose_files, width=30)
        self.entry_spec2.grid(row=1, column=1, padx=5, pady=5)

        self.combo_column1 = ttk.Combobox(self.frame_choose_columns, textvariable=self.column1_var, state='readonly')
        self.combo_column1.grid(row=0, column=1, padx=5, pady=5)

        self.combo_column2 = ttk.Combobox(self.frame_choose_columns, textvariable=self.column2_var, state='readonly')
        self.combo_column2.grid(row=1, column=1, padx=5, pady=5)

        self.choose_files_button = tk.Button(self.frame_choose_files, text="Choisir les fichiers", command=self.choose_files)
        self.choose_files_button.grid(row=0, column=2, padx=5, pady=5, rowspan=2)

        self.compare_button = tk.Button(self.frame_buttons, text="Comparer", command=self.compare_excel_files)
        self.compare_button.grid(row=0, column=0, padx=5, pady=5)

        self.create_log_button = tk.Button(self.frame_buttons, text="Créer fichier log", command=self.create_log_file)
        self.create_log_button.grid(row=0, column=1, padx=5, pady=5)

        self.export_excel_button = tk.Button(self.frame_buttons, text="Exporter vers Excel", command=self.export_to_excel)
        self.export_excel_button.grid(row=0, column=2, padx=5, pady=5)

        self.output_tree = ttk.Treeview(self.frame_output, columns=('Index', 'Value from File 1', 'Value from File 2', 'Différence'))
        self.output_tree.heading('#0', text='Index')
        self.output_tree.heading('Index', text='Index')
        self.output_tree.heading('Value from File 1', text='Value from File 1')
        self.output_tree.heading('Value from File 2', text='Value from File 2')
        self.output_tree.heading('Différence', text='Différence')
        self.output_tree.column('#0', width=50)
        self.output_tree.column('Index', width=50, anchor='center')
        self.output_tree.column('Value from File 1', width=200, anchor='center')
        self.output_tree.column('Value from File 2', width=200, anchor='center')
        self.output_tree.column('Différence', width=100, anchor='center')

        vsb = ttk.Scrollbar(self.frame_output, orient="vertical", command=self.output_tree.yview)
        vsb.pack(side='right', fill='y')
        self.output_tree.configure(yscrollcommand=vsb.set)

        self.output_tree.pack(expand=True, fill=tk.BOTH)

        self.column1_var.set("Column Name")
        self.column2_var.set("Column Name")

    def choose_files(self):
        file1 = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx")])
        file2 = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx")])

        if file1 and file2:
            self.entry_spec1.delete(0, tk.END)
            self.entry_spec1.insert(0, file1)

            self.entry_spec2.delete(0, tk.END)
            self.entry_spec2.insert(0, file2)

            self.update_column_comboboxes(file1, file2)

    def update_column_comboboxes(self, file1, file2):
        try:
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)

            columns1 = df1.columns.tolist()
            columns2 = df2.columns.tolist()

            self.combo_column1['values'] = columns1
            self.combo_column2['values'] = columns2

            if 'Column Name' in columns1:
                self.column1_var.set('Column Name')
            if 'Column Name' in columns2:
                self.column2_var.set('Column Name')

        except pd.errors.EmptyDataError:
            messagebox.showerror("Erreur", "Le fichier Excel est vide ou ne contient pas de données.")
        except FileNotFoundError:
            messagebox.showerror("Erreur", "Fichier non trouvé. Vérifiez le chemin du fichier.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")

    def compare_excel_files(self):
        spec1_file = self.entry_spec1.get()
        spec2_file = self.entry_spec2.get()
        column1 = self.column1_var.get()
        column2 = self.column2_var.get()

        try:
            df1 = pd.read_excel(spec1_file)
            df2 = pd.read_excel(spec2_file)

            if column1 in df1.columns and column2 in df2.columns:
                diff_values = []
                for value1, value2 in zip(df1[column1], df2[column2]):
                    if value1 == value2:
                        diff_values.append('')
                    else:
                        diff_values.append(f'{value1} ≠ {value2}')

                self.output_tree.delete(*self.output_tree.get_children())

                for index, (value1, value2, diff) in enumerate(zip(df1[column1], df2[column2], diff_values), start=1):
                    self.output_tree.insert('', tk.END, text=index, values=(index, value1, value2, diff))

                self.comparison_data = list(zip(range(1, len(diff_values) + 1), df1[column1], df2[column2], diff_values))

            else:
                messagebox.showerror("Erreur", "Veuillez sélectionner des colonnes valides pour la comparaison.")

        except pd.errors.EmptyDataError:
            messagebox.showerror("Erreur", "Le fichier Excel est vide ou ne contient pas de données.")
        except FileNotFoundError:
            messagebox.showerror("Erreur", "Fichier non trouvé. Vérifiez le chemin du fichier.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")

    def create_log_file(self):
        try:
            if not hasattr(self, 'comparison_data'):
                messagebox.showerror("Erreur", "Aucune comparaison n'a été effectuée. Veuillez comparer d'abord les fichiers Excel.")
                return

            file_path = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Fichiers log", "*.log")])
            if file_path:
                with open(file_path, 'w') as file:
                    file.write("Index\tValue from File 1\tValue from File 2\tDifférence\n")
                    for item in self.comparison_data:
                        file.write(f"{item[0]}\t{item[1]}\t{item[2]}\t{item[3]}\n")

                messagebox.showinfo("Fichier Log créé", f"Le fichier log a été créé avec succès : {file_path}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite lors de la création du fichier log : {str(e)}")

    def export_to_excel(self):
        try:
            if not hasattr(self, 'comparison_data'):
                messagebox.showerror("Erreur", "Aucune comparaison n'a été effectuée. Veuillez comparer d'abord les fichiers Excel.")
                return

            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Fichiers Excel", "*.xlsx")])
            if file_path:
                df = pd.DataFrame(self.comparison_data, columns=['Index', 'Value from File 1', 'Value from File 2', 'Différence'])
                df.to_excel(file_path, index=False)

                messagebox.showinfo("Export Excel réussi", f"Les données de comparaison ont été exportées avec succès vers : {file_path}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite lors de l'exportation vers Excel : {str(e)}")


def main():
    root = tk.Tk()
    app = ExcelComparatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
