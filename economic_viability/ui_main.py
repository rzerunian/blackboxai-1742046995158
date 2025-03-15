import tkinter as tk
from tkinter import ttk, messagebox, filedialog

class MainApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Análise de Viabilidade Econômica")
        self.root.geometry("1200x800")
        
        # Apply theme
        style = ttk.Style()
        style.theme_use('clam')
        
        # Create main notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Initialize frames
        self.project_frame = self.create_project_frame()
        self.capex_frame = self.create_capex_frame()
        self.opex_frame = self.create_opex_frame()
        self.receitas_frame = self.create_receitas_frame()
        self.config_frame = self.create_config_frame()
        self.analysis_frame = self.create_analysis_frame()
        
        # Add frames to notebook
        self.notebook.add(self.project_frame, text='Projeto')
        self.notebook.add(self.capex_frame, text='CapEx')
        self.notebook.add(self.opex_frame, text='OpEx')
        self.notebook.add(self.receitas_frame, text='Receitas')
        self.notebook.add(self.config_frame, text='Configuração')
        self.notebook.add(self.analysis_frame, text='Análise')

    def create_project_frame(self):
        frame = ttk.Frame(self.notebook)
        
        # Project registration section
        reg_frame = ttk.LabelFrame(frame, text="Cadastro de Projeto", padding="10")
        reg_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(reg_frame, text="Nome do Projeto:").grid(row=0, column=0, padx=5, pady=5)
        self.project_name = ttk.Entry(reg_frame, width=40)
        self.project_name.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(reg_frame, text="Descrição:").grid(row=1, column=0, padx=5, pady=5)
        self.project_desc = ttk.Entry(reg_frame, width=40)
        self.project_desc.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Button(reg_frame, text="Cadastrar Projeto", command=self.register_project).grid(row=2, column=1, pady=10)
        
        # Project selection section
        sel_frame = ttk.LabelFrame(frame, text="Seleção de Projeto", padding="10")
        sel_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(sel_frame, text="Projeto:").grid(row=0, column=0, padx=5, pady=5)
        self.project_select = ttk.Combobox(sel_frame, width=37)
        self.project_select.grid(row=0, column=1, padx=5, pady=5)
        
        return frame

    def create_capex_frame(self):
        frame = ttk.Frame(self.notebook)
        
        # Controls frame
        controls = ttk.Frame(frame)
        controls.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(controls, text="Adicionar Item", command=self.add_capex).pack(side='left', padx=5)
        ttk.Button(controls, text="Importar Excel", command=self.import_capex).pack(side='left', padx=5)
        
        # Treeview
        columns = ('tag', 'descricao', 'quantidade', 'valor_unitario', 'valor_total')
        self.capex_tree = ttk.Treeview(frame, columns=columns, show='headings')
        
        # Define headings
        self.capex_tree.heading('tag', text='TAG')
        self.capex_tree.heading('descricao', text='Descrição')
        self.capex_tree.heading('quantidade', text='Quantidade')
        self.capex_tree.heading('valor_unitario', text='Valor Unitário')
        self.capex_tree.heading('valor_total', text='Valor Total')
        
        # Define columns
        self.capex_tree.column('tag', width=100)
        self.capex_tree.column('descricao', width=300)
        self.capex_tree.column('quantidade', width=100)
        self.capex_tree.column('valor_unitario', width=100)
        self.capex_tree.column('valor_total', width=100)
        
        self.capex_tree.pack(expand=True, fill='both', padx=10, pady=5)
        
        return frame

    def create_opex_frame(self):
        frame = ttk.Frame(self.notebook)
        
        # Controls frame
        controls = ttk.Frame(frame)
        controls.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(controls, text="Adicionar Item", command=self.add_opex).pack(side='left', padx=5)
        ttk.Button(controls, text="Importar Excel", command=self.import_opex).pack(side='left', padx=5)
        
        # Treeview
        columns = ('tag', 'descricao', 'quantidade', 'valor_unitario', 'valor_total')
        self.opex_tree = ttk.Treeview(frame, columns=columns, show='headings')
        
        # Define headings
        self.opex_tree.heading('tag', text='TAG')
        self.opex_tree.heading('descricao', text='Descrição')
        self.opex_tree.heading('quantidade', text='Quantidade')
        self.opex_tree.heading('valor_unitario', text='Valor Unitário')
        self.opex_tree.heading('valor_total', text='Valor Total')
        
        # Define columns
        self.opex_tree.column('tag', width=100)
        self.opex_tree.column('descricao', width=300)
        self.opex_tree.column('quantidade', width=100)
        self.opex_tree.column('valor_unitario', width=100)
        self.opex_tree.column('valor_total', width=100)
        
        self.opex_tree.pack(expand=True, fill='both', padx=10, pady=5)
        
        return frame

    def create_receitas_frame(self):
        frame = ttk.Frame(self.notebook)
        
        # Controls frame
        controls = ttk.Frame(frame)
        controls.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(controls, text="Adicionar Item", command=self.add_receita).pack(side='left', padx=5)
        ttk.Button(controls, text="Importar Excel", command=self.import_receita).pack(side='left', padx=5)
        
        # Treeview
        columns = ('tag', 'descricao', 'quantidade', 'valor_unitario', 'valor_total')
        self.receitas_tree = ttk.Treeview(frame, columns=columns, show='headings')
        
        # Define headings
        self.receitas_tree.heading('tag', text='TAG')
        self.receitas_tree.heading('descricao', text='Descrição')
        self.receitas_tree.heading('quantidade', text='Quantidade')
        self.receitas_tree.heading('valor_unitario', text='Valor Unitário')
        self.receitas_tree.heading('valor_total', text='Valor Total')
        
        # Define columns
        self.receitas_tree.column('tag', width=100)
        self.receitas_tree.column('descricao', width=300)
        self.receitas_tree.column('quantidade', width=100)
        self.receitas_tree.column('valor_unitario', width=100)
        self.receitas_tree.column('valor_total', width=100)
        
        self.receitas_tree.pack(expand=True, fill='both', padx=10, pady=5)
        
        return frame

    def create_config_frame(self):
        frame = ttk.Frame(self.notebook)
        
        # Tax configuration
        tax_frame = ttk.LabelFrame(frame, text="Configurações Tributárias", padding="10")
        tax_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(tax_frame, text="Alíquota IR (%):").grid(row=0, column=0, padx=5, pady=5)
        self.ir_rate = ttk.Entry(tax_frame, width=10)
        self.ir_rate.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(tax_frame, text="Alíquota CSLL (%):").grid(row=1, column=0, padx=5, pady=5)
        self.csll_rate = ttk.Entry(tax_frame, width=10)
        self.csll_rate.grid(row=1, column=1, padx=5, pady=5)
        
        # TMA configuration
        tma_frame = ttk.LabelFrame(frame, text="Taxa Mínima de Atratividade", padding="10")
        tma_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(tma_frame, text="TMA (%):").grid(row=0, column=0, padx=5, pady=5)
        self.tma_rate = ttk.Entry(tma_frame, width=10)
        self.tma_rate.grid(row=0, column=1, padx=5, pady=5)
        
        # Save button
        ttk.Button(frame, text="Salvar Configurações", command=self.save_config).pack(pady=10)
        
        return frame

    def create_analysis_frame(self):
        frame = ttk.Frame(self.notebook)
        
        # Results frame
        results_frame = ttk.LabelFrame(frame, text="Resultados da Análise", padding="10")
        results_frame.pack(fill='x', padx=10, pady=5)
        
        # TIR
        ttk.Label(results_frame, text="TIR:").grid(row=0, column=0, padx=5, pady=5)
        self.tir_result = ttk.Label(results_frame, text="--")
        self.tir_result.grid(row=0, column=1, padx=5, pady=5)
        
        # VPL
        ttk.Label(results_frame, text="VPL:").grid(row=1, column=0, padx=5, pady=5)
        self.vpl_result = ttk.Label(results_frame, text="--")
        self.vpl_result.grid(row=1, column=1, padx=5, pady=5)
        
        # Payback
        ttk.Label(results_frame, text="Payback:").grid(row=2, column=0, padx=5, pady=5)
        self.payback_result = ttk.Label(results_frame, text="--")
        self.payback_result.grid(row=2, column=1, padx=5, pady=5)
        
        # Dívida Líquida/EBITDA
        ttk.Label(results_frame, text="Dívida Líquida/EBITDA:").grid(row=3, column=0, padx=5, pady=5)
        self.debt_ebitda_result = ttk.Label(results_frame, text="--")
        self.debt_ebitda_result.grid(row=3, column=1, padx=5, pady=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Calcular", command=self.calculate_analysis).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Exportar para Excel", command=self.export_analysis).pack(side='left', padx=5)
        
        return frame

    # Callback methods (to be implemented)
    def register_project(self):
        pass

    def add_capex(self):
        pass

    def import_capex(self):
        pass

    def add_opex(self):
        pass

    def import_opex(self):
        pass

    def add_receita(self):
        pass

    def import_receita(self):
        pass

    def save_config(self):
        pass

    def calculate_analysis(self):
        pass

    def export_analysis(self):
        pass
