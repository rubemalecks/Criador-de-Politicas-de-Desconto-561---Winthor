"""Gerador de políticas WinThor 561 — V12.2: OCI empacotado (Thick Mode); identificação do usuário Oracle; placeholders com TAB e validação obrigatória; PERCDESC com vírgula e 5 casas.
Executar: python gerador_561_v12.py (requer python-oracledb para testar conexão).
"""
import calendar
import csv
import sys
import os
import threading
import re

# Importação explícita para que o PyInstaller detecte o driver Oracle.
try:
    import oracledb
    ORACLE_IMPORT_ERROR = None
except ImportError as exc:
    oracledb = None
    ORACLE_IMPORT_ERROR = (type(exc).__name__, getattr(exc, "name", None))

import tkinter as tk
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

HEADERS = ['Cód.Prod.','Início Vigência','Fim Vigência','Percentual Com.','Cód.Fornec.','Cód.Depto.','Cód.Seção','Cód.Categoria','Região','Cód.Praça','Cód.Superv.','Cód.RCA','Cód.Cli.','Cód.Ramo Ativ.','Cód.Plano Pagto.','Alterar Pr.Tabela','Aplicar Desconto','Base Déb.Créd.RCA','Classe Venda','Cód.Prod.Cesta','Cód.Prod.Princ.','Credita sobre Política','Origem Ped.','Tipo Aplica Desc','Tipo Carga','Utiliza Desc.Rede','Qtde Inicial','Qtde Final','Prioritária','Prioritária Geral','Descrição','Verba','% Custeado pelo Fornec.','Cód. Filial']

def dec(s):
    s = str(s).strip().replace('R$', '').replace(' ', '')
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    return Decimal(s)

def integer(s, name):
    if not str(s).strip():
        raise ValueError(f'{name}: informe um valor.')
    try:
        v = dec(s)
    except InvalidOperation:
        raise ValueError(f'{name}: informe um número válido.') from None
    if v != v.to_integral_value() or v <= 0:
        raise ValueError(f'{name} deve ser inteiro positivo.')
    return int(v)

def date(s):
    d = datetime.strptime(s.strip(), '%d/%m/%Y')
    return d

DEFAULTS = {15:'S',16:'S',17:'S',21:'N',22:'O',23:'A',25:'N',28:'S',29:'S'}
# Campos que aceitam S/N, códigos inteiros, números decimais ou texto.
SN_FIELDS = {15,16,17,21,25,28,29}
DECIMAL_FIELDS = {3,31,32}
TEXT_FIELDS = {18,22,23,24,30}
DATE_FIELDS = {1,2}
AUTO_FIELDS = {0,1,2,3,11,20,26,27,30}

def validate_extra(index, value):
    value = value.strip()
    if not value: return ''
    if any(c in value for c in ('\t','\r','\n')):
        raise ValueError(f'{HEADERS[index]}: não utilize tabulações ou quebras de linha.')
    if index in SN_FIELDS:
        if value not in ('S','N'): raise ValueError(f'{HEADERS[index]}: selecione S ou N.')
    elif index in DATE_FIELDS:
        date(value)
    elif index in DECIMAL_FIELDS:
        try: dec(value)
        except InvalidOperation: raise ValueError(f'{HEADERS[index]}: informe um número decimal.')
    elif index not in TEXT_FIELDS:
        if not value.isdigit(): raise ValueError(f'{HEADERS[index]}: informe somente números inteiros.')
    return value

def make_rows(data, ranges, extras=None):
    code = str(data['codigo']).strip()
    if not code.isdigit() or int(code) <= 0:
        raise ValueError('Informe um código de produto válido.')
    name = data['nome'].strip()
    if not name or '\t' in name or '\n' in name or '\r' in name:
        raise ValueError('Informe nome de campanha válido, sem tabulação ou quebra de linha.')
    unit = integer(data['unidade'], 'Unidade de venda')
    base = dec(data['tabela'])
    if base <= 0: raise ValueError('Preço de tabela deve ser maior que zero.')
    start, end = date(data['inicio']), date(data['fim'])
    if end < start: raise ValueError('Data final anterior à inicial.')
    if not ranges: raise ValueError('Cadastre ao menos uma faixa.')
    if not str(data['tabela']).strip(): raise ValueError('Preencha o preço de tabela.')
    rca = '' if data.get('todos_rca') else validate_extra(11, data.get('rca', ''))
    extras = extras or {}
    validated = {idx:validate_extra(idx,value) for idx,value in extras.items()}
    rows, previous = [], 0
    for i, (start_qty, end_qty, new_price) in enumerate(ranges, 1):
        low, high = integer(start_qty, f'Quantidade inicial da faixa {i}'), integer(end_qty, f'Quantidade final da faixa {i}')
        if high < low: raise ValueError(f'Faixa {i}: quantidade final menor que a inicial.')
        if low <= previous: raise ValueError(f'Faixa {i}: quantidades sobrepostas ou fora de ordem.')
        previous = high
        if not str(new_price).strip() or dec(new_price) <= 0:
            raise ValueError(f'Faixa {i}: informe um preço novo maior que zero.')
        price = dec(new_price)
        if price < 0 or price > base: raise ValueError(f'Faixa {i}: preço novo deve estar entre zero e preço de tabela.')
        percent = (base - price) / base * 100
        row = [''] * 34
        row[0 if data['tipo'] == 'produto' else 20] = code
        row[1], row[2] = start.strftime('%d/%m/%Y'), end.strftime('%d/%m/%Y')
        row[3] = format(percent.quantize(Decimal('0.00001'), rounding=ROUND_HALF_UP), '.5f').replace('.', ',')
        row[11] = rca
        for idx,value in DEFAULTS.items(): row[idx] = value
        for idx,value in validated.items(): row[idx] = value
        row[26], row[27] = str(low * unit), str(high * unit)
        row[30] = name
        rows.append(row)
    return rows

class RoundedEntry(tk.Frame):
    """Campo compacto com fundo e contorno arredondados, sem dependências."""
    def __init__(self, parent, textvariable, width=150, numeric=False):
        super().__init__(parent, bg='#101114', width=width, height=30)
        self.pack_propagate(False)
        canvas=tk.Canvas(self, width=width, height=30, bg='#101114',
                         highlightthickness=0, bd=0)
        canvas.place(x=0,y=0)
        def roundrect(x1,y1,x2,y2,r,fill,outline):
            pts=[x1+r,y1,x2-r,y1,x2,y1,x2,y1+r,x2,y2-r,x2,y2,x2-r,y2,
                 x1+r,y2,x1,y2,x1,y2-r,x1,y1+r,x1,y1]
            canvas.create_polygon(pts,smooth=True,splinesteps=12,
                                  fill=fill,outline=outline,width=1)
        roundrect(1,1,width-1,29,7,'#242832','#454b59')
        self.entry=tk.Entry(self,textvariable=textvariable,bg='#242832',fg='#f0f2f6',
                            insertbackground='white',selectbackground='#365a91',
                            disabledbackground='#30343c',disabledforeground='#a0a7b3',
                            readonlybackground='#30343c',
                            relief='flat',bd=0,font=('Segoe UI',10))
        self.entry.place(x=10,y=6,width=width-20,height=19)
        if numeric:
            cmd=self.register(lambda value: value=='' or value.isdigit())
            self.entry.configure(validate='key',validatecommand=(cmd,'%P'))

class PlaceholderEntry(RoundedEntry):
    """Exibe sugestão sem inseri-la no valor; TAB aceita o valor sugerido."""
    def __init__(self, parent, textvariable, placeholder='', width=150, numeric=False):
        super().__init__(parent, textvariable, width=width, numeric=numeric)
        self.placeholder = str(placeholder)
        self.variable = textvariable
        self.hint = tk.Label(self, text=self.placeholder, fg='#8793a8', bg='#242832',
                             font=('Segoe UI', 10), anchor='w', cursor='xterm')
        self.hint.bind('<Button-1>', self.activate)
        self.entry.bind('<FocusIn>', lambda e: self.hint.place_forget(), add='+')
        self.entry.bind('<FocusOut>', lambda e: self.after_idle(self.refresh), add='+')
        self.entry.bind('<Tab>', self.accept_tab, add='+')
        self.entry.bind('<ISO_Left_Tab>', self.accept_tab, add='+')
        self.entry.bind('<Shift-Tab>', self.accept_tab, add='+')
        self.variable.trace_add('write', lambda *_: self.refresh())
        self.after_idle(self.refresh)
    def activate(self, event=None):
        self.hint.place_forget()
        self.entry.focus_set()
    def refresh(self):
        if not self.winfo_exists(): return
        if not self.variable.get() and self.entry.focus_get() != self.entry:
            self.hint.place(x=10, y=5, width=self.winfo_reqwidth()-20, height=20)
        else:
            self.hint.place_forget()
    def accept_tab(self, event):
        if not self.variable.get().strip() and self.placeholder:
            self.variable.set(self.placeholder)
        # Deixa o Tkinter continuar a navegação padrão por TAB.

class RoundedButton(tk.Canvas):
    """Botão com cantos arredondados e suporte a desabilitado."""
    def __init__(self, parent, text, command, width=110, height=30):
        super().__init__(parent, width=width, height=height, bg='#101114',
                         highlightthickness=0, bd=0, cursor='hand2')
        self.command=command
        self.enabled=True
        self.label=text
        self.w=width; self.h=height
        self.bind('<Button-1>',self.clicked)
        self.bind('<Enter>',lambda e:self.draw(hover=True))
        self.bind('<Leave>',lambda e:self.draw())
        self.draw()
    def draw(self,hover=False):
        self.delete('all')
        color='#354560' if hover and self.enabled else '#252b36'
        border='#6b8cc3' if hover and self.enabled else '#4b5565'
        if not self.enabled: color='#1a1c22'; border='#343841'
        r=8; w=self.w; h=self.h
        self.create_polygon(r,1,w-r,1,w-1,1,w-1,r,w-1,h-r,w-1,h-1,
                            w-r,h-1,r,h-1,1,h-1,1,h-r,1,r,1,1,
                            smooth=True,splinesteps=12,fill=color,outline=border)
        self.create_text(w/2,h/2,text=self.label,fill='#f0f2f6' if self.enabled else '#727986',font=('Segoe UI',9))
        self.configure(cursor='hand2' if self.enabled else 'arrow')
    def clicked(self,event):
        if self.enabled:self.command()
    def configure(self, cnf=None, **kw):
        if 'state' in kw:
            self.enabled=kw.pop('state')!='disabled'
            self.draw()
        if kw: super().configure(cnf,**kw)

class RoundedChoice(tk.Frame):
    """Seletor S/N com bordas arredondadas; aceita vazio."""
    def __init__(self,parent,variable,choices,width=175):
        super().__init__(parent,bg='#101114',width=width,height=30)
        self.grid_propagate(False)
        self.variable=variable
        self.choices=choices
        self.button=RoundedButton(self,self.text(),self.advance,width=width,height=30)
        self.button.place(x=0,y=0)
        variable.trace_add('write',lambda *_:self.refresh())
    def text(self):return (self.variable.get() or '— vazio —')+'  ▾'
    def refresh(self):self.button.label=self.text();self.button.draw()
    def advance(self):
        try: index=self.choices.index(self.variable.get())
        except ValueError: index=-1
        self.variable.set(self.choices[(index+1)%len(self.choices)])

def winthor_context(argv=None):
    """Lê apenas metadados não secretos dos 5 argumentos da ação WinThor.

    Ordem: USUARIOWT, SENHABD, ALIASBD, USUARIOBD, CODROTINA.
    A senha (posição 2) nunca é exibida, salva ou registrada.
    """
    args = sys.argv[1:] if argv is None else argv
    if len(args) != 5:
        return None
    return {
        'usuario': args[0].strip(),
        'senha_bd': args[1],
        'usuario_bd': args[3].strip(),
        'alias': args[2].strip(),
        'rotina': args[4].strip(),
    }



def find_tns_directory(alias):
    """Localiza tnsnames.ora por configuração, sem depender da versão Oracle.

    Prioridade: TNS_ADMIN, ORACLE_HOME, Oracle homes no registro Windows,
    diretórios do PATH e instalações comuns. Exige o alias solicitado.
    """
    if not re.fullmatch(r"[A-Za-z0-9_.-]+", alias):
        return None
    candidates = []
    seen = set()

    def add(directory):
        if not directory:
            return
        try:
            path = Path(os.path.expandvars(str(directory).strip('" ')))
            key = str(path).casefold()
            if key not in seen:
                seen.add(key)
                candidates.append(path)
        except (OSError, ValueError):
            pass

    add(os.environ.get('TNS_ADMIN'))
    home = os.environ.get('ORACLE_HOME')
    if home:
        add(Path(home) / 'network' / 'admin')

    if os.name == 'nt':
        try:
            import winreg
            for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for base in (r'SOFTWARE\ORACLE', r'SOFTWARE\WOW6432Node\ORACLE'):
                    try:
                        with winreg.OpenKey(root, base) as key:
                            keys = ['']
                            index = 0
                            while True:
                                try:
                                    keys.append(winreg.EnumKey(key, index))
                                    index += 1
                                except OSError:
                                    break
                        for sub in keys:
                            try:
                                with winreg.OpenKey(root, base + ('\\' + sub if sub else '')) as item:
                                    try:
                                        add(winreg.QueryValueEx(item, 'TNS_ADMIN')[0])
                                    except OSError:
                                        pass
                                    try:
                                        add(Path(winreg.QueryValueEx(item, 'ORACLE_HOME')[0]) / 'network' / 'admin')
                                    except OSError:
                                        pass
                            except OSError:
                                pass
                    except OSError:
                        pass
        except ImportError:
            pass

    for entry in os.environ.get('PATH', '').split(os.pathsep):
        if 'oracle' in entry.lower() or 'client' in entry.lower():
            path = Path(entry.strip('" '))
            add(path / 'network' / 'admin')
            add(path.parent / 'network' / 'admin')

    for root in (Path('C:/oracle/product'), Path('C:/app'), Path('C:/Oracle')):
        if not root.is_dir():
            continue
        # Busca limitada: instalações Oracle normalmente têm network/admin
        # até alguns níveis abaixo da pasta de instalação.
        folders = [root]
        for _ in range(4):
            next_folders = []
            for folder in folders:
                add(folder / 'network' / 'admin')
                try:
                    next_folders.extend(x for x in folder.iterdir() if x.is_dir() and x.name.lower() != 'network')
                except (OSError, PermissionError):
                    pass
            folders = next_folders[:300]

    # Preferir o primeiro arquivo que contém definição do alias pedido.
    # Evita depender de uma versão fixa do Oracle Client.
    pattern = re.compile(r'(?im)^\s*' + re.escape(alias) + r'\s*=')
    for folder in candidates:
        file = folder / 'tnsnames.ora'
        try:
            if file.is_file() and pattern.search(file.read_text(encoding='cp1252', errors='replace')):
                return str(folder)
        except (OSError, UnicodeError):
            continue
    return None

def bundled_oracle_client():
    """Pasta OCI distribuída com o EXE (PyInstaller --onefile)."""
    base = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    folder = base / 'oracle_client'
    return folder if (folder / 'oci.dll').is_file() else None

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.winthor = winthor_context()
        self.title('Gerador de Politicas de Desconto - Rotina 561')
        self.geometry('750x650')
        self.minsize(630, 470)
        self.configure(bg='#101114')
        style = ttk.Style(self)
        if 'clam' in style.theme_names(): style.theme_use('clam')
        bg='#101114'; panel='#191c22'; inputbg='#242832'; fg='#f0f2f6'; muted='#aab2c0'; accent='#60a5fa'
        style.configure('TFrame', background=bg)
        style.configure('TLabel', background=bg, foreground=fg)
        style.configure('TButton', padding=8, background=inputbg, foreground=fg, bordercolor='#414857')
        style.map('TButton', background=[('active','#364256'),('pressed','#233c63')],foreground=[('disabled','#777'),('active','#ffffff')])
        style.configure('TEntry', fieldbackground=inputbg, foreground=fg, insertcolor=fg, bordercolor='#454b59', padding=5)
        style.configure('TCombobox', fieldbackground=inputbg, background=inputbg, foreground=fg, arrowcolor=fg, padding=5, bordercolor='#454b59', lightcolor=inputbg, darkcolor=inputbg, selectbackground=inputbg, selectforeground=fg)
        style.map('TCombobox', fieldbackground=[('readonly',inputbg),('disabled',panel)], background=[('readonly',inputbg),('active','#303744')], foreground=[('readonly',fg),('disabled',muted)], selectbackground=[('readonly',inputbg)], selectforeground=[('readonly',fg)])
        style.configure('TCheckbutton', background=bg, foreground=fg, font=('Segoe UI', 9), padding=2)
        style.map('TCheckbutton', background=[('active',bg),('disabled',bg)], foreground=[('active',fg),('disabled',muted)])
        style.configure('Dark.TCheckbutton', background=bg, foreground=fg, padding=2)
        style.map('Dark.TCheckbutton', background=[('active',bg)], foreground=[('active',fg)])
        style.configure('TRadiobutton',background=bg,foreground=fg)
        style.map('TRadiobutton',background=[('active',bg)],foreground=[('active',fg)])
        style.configure('TNotebook',background=bg,borderwidth=0)
        style.configure('TNotebook.Tab',background=panel,foreground=muted,padding=(18,8),borderwidth=0)
        style.map('TNotebook.Tab',background=[('selected','#303b50'),('active','#28303e')],foreground=[('selected','#ffffff'),('active','#ffffff')],padding=[('selected',(18,9)),('!selected',(18,9))])
        style.configure('Vertical.TScrollbar',background=inputbg,troughcolor=bg,arrowcolor=fg)
        self.option_add('*TCombobox*Listbox.background', inputbg)
        self.option_add('*TCombobox*Listbox.foreground', fg)
        self.option_add('*TCombobox*Listbox.selectBackground', '#365a91')
        self.vars = {key: tk.StringVar(value=val) for key, val in {'tipo':'produto','codigo':'','nome':'','unidade':'','tabela':'','inicio':datetime.today().strftime('%d/%m/%Y'),'fim':datetime.today().strftime('%d/%m/%Y'),'rca':''}.items()}
        self.ranges = []
        self.todos_rca = tk.BooleanVar(value=False)
        self.infinita = tk.BooleanVar(value=False)
        self.regiao = tk.StringVar()
        self.regioes = {}
        self.qtunits = {}
        self._oracle_lock = threading.Lock()
        self._selected_product = None
        self._price_raw = None
        self._load_token = 0
        self._product_search_token = 0
        self._results_popup = None
        self._rca_search_token = 0
        self.db_status = tk.StringVar(value="Oracle: não testado")
        main = ttk.Frame(self, padding=12)
        main.pack(fill='both', expand=True)
        ttk.Label(main, text='Gerador de Politicas de Desconto - Rotina 561', font=('Segoe UI', 15, 'bold')).pack(anchor='w')
        ttk.Label(main, text='HOMOLOGAÇÃO • RCA vazio por padrão • TXT separado por TAB').pack(anchor='w', pady=(1,2))
        if self.winthor:
            usuario = self.winthor['usuario'] or '(não informado)'
            alias = self.winthor['alias'] or '(não informado)'
            usuario_bd = self.winthor['usuario_bd'] or '(não informado)'
            ttk.Label(main, text=f'WinThor: {usuario}   •   Oracle: {usuario_bd}   •   TNS: {alias}',
                      foreground='#aab2c0').pack(anchor='w', pady=(0,6))
        else:
            ttk.Label(main, text='Execução direta pelo Windows (sem parâmetros WinThor)',
                      foreground='#aab2c0').pack(anchor='w', pady=(0,6))
        # Rodapé fixo: os botões permanecem visíveis em qualquer tamanho da janela.
        footer=ttk.Frame(main)
        footer.pack(side='bottom',fill='x',pady=(8,0))
        controls=ttk.Frame(footer); controls.pack(fill='x')
        RoundedButton(controls,text='Pré-visualizar TXT',command=self.preview,width=145).pack(side='left',padx=(0,8))
        RoundedButton(controls,text='Gerar TXT para importar',command=self.save,width=174).pack(side='left')
        self.db_button = RoundedButton(controls,text='Testar conexão Oracle',command=self.test_oracle,width=175)
        self.db_button.pack(side='left',padx=(8,0))
        ttk.Label(footer,textvariable=self.db_status,foreground='#aab2c0').pack(anchor='w',pady=(4,0))
        ttk.Label(footer,text='Prévia do TXT (34 colunas separadas por TAB):',foreground='#aab2c0').pack(anchor='w',pady=(7,3))
        self.output=tk.Text(footer,height=2,wrap='none',font=('Consolas',9),bg='#191c22',fg='#f0f2f6',insertbackground='#f0f2f6',selectbackground='#365a91',relief='flat',highlightthickness=1,highlightbackground='#414857')
        self.output.pack(fill='x')
        tabs=ttk.Notebook(main); tabs.pack(fill='both',expand=True)
        basic_tab=ttk.Frame(tabs,padding=0); detail=ttk.Frame(tabs,padding=8)
        tabs.add(basic_tab,text='Política resumida'); tabs.add(detail,text='Outros detalhes')
        # Área principal rolável: faixas e botão de adicionar não desaparecem em janelas pequenas.
        basic_canvas=tk.Canvas(basic_tab,bg='#101114',highlightthickness=0)
        basic_scroll=ttk.Scrollbar(basic_tab,orient='vertical',command=basic_canvas.yview)
        basic_canvas.configure(yscrollcommand=basic_scroll.set)
        basic_canvas.pack(side='left',fill='both',expand=True)
        def update_scrollbar(*_):
            box=basic_canvas.bbox('all')
            overflow=bool(box and box[3]>basic_canvas.winfo_height()+2)
            if overflow and not basic_scroll.winfo_manager():
                basic_scroll.pack(side='right',fill='y')
            elif not overflow and basic_scroll.winfo_manager():
                basic_scroll.pack_forget()
                basic_canvas.yview_moveto(0)
        self.after_idle(update_scrollbar)
        basic=ttk.Frame(basic_canvas,padding=(7,5,7,7))
        basic_window=basic_canvas.create_window((0,0),window=basic,anchor='nw')
        def basic_content_resized(event):
            basic_canvas.configure(scrollregion=basic_canvas.bbox('all'))
            self.after_idle(update_scrollbar)
        basic.bind('<Configure>',basic_content_resized)
        def basic_canvas_resized(event):
            basic_canvas.itemconfigure(basic_window,width=event.width)
            self.after_idle(update_scrollbar)
        basic_canvas.bind('<Configure>',basic_canvas_resized)
        def scroll_basic(event):
            if basic_canvas.bbox('all') and basic_canvas.bbox('all')[3] > basic_canvas.winfo_height():
                basic_canvas.yview_scroll(-1 if event.delta>0 else 1,'units')
        def bind_wheel(event):
            basic_canvas.bind_all('<MouseWheel>',scroll_basic)
        def unbind_wheel(event):
            basic_canvas.unbind_all('<MouseWheel>')
        basic_canvas.bind('<Enter>',bind_wheel)
        basic_canvas.bind('<Leave>',unbind_wheel)
        self.extra_vars={}
        detail_canvas=tk.Canvas(detail,highlightthickness=0,bg='#101114')
        detail_scroll=ttk.Scrollbar(detail,orient='vertical',command=detail_canvas.yview)
        detail_body=ttk.Frame(detail_canvas)
        detail_body.bind('<Configure>',lambda e:detail_canvas.configure(scrollregion=detail_canvas.bbox('all')))
        detail_canvas.create_window((0,0),window=detail_body,anchor='nw')
        detail_canvas.configure(yscrollcommand=detail_scroll.set)
        detail_canvas.pack(side='left',fill='both',expand=True)
        detail_scroll.pack(side='right',fill='y')
        ttk.Label(detail_body,text='Campos adicionais do TXT (vazios permanecem vazios).',font=('Segoe UI',10,'bold')).grid(row=0,column=0,columnspan=2,sticky='w',pady=(0,10))
        extra_indices=[i for i in range(34) if i not in AUTO_FIELDS]
        for n,idx in enumerate(extra_indices):
            v=tk.StringVar(value=DEFAULTS.get(idx,''));self.extra_vars[idx]=v
            r=n+1
            ttk.Label(detail_body,text=f'{idx+1:02d} — {HEADERS[idx]}').grid(row=r,column=0,sticky='w',padx=(0,18),pady=4)
            if idx in SN_FIELDS:
                widget=RoundedChoice(detail_body,variable=v,choices=['','S','N'],width=175)
            elif idx==23:
                widget=RoundedChoice(detail_body,variable=v,choices=['','A'],width=175)
            else:
                widget=RoundedEntry(detail_body,textvariable=v,width=175)
                if idx not in TEXT_FIELDS and idx not in DATE_FIELDS:
                    pattern=(lambda proposed,ix=idx: proposed=='' or (all(c.isdigit() for c in proposed) if ix not in DECIMAL_FIELDS else all(c.isdigit() or c in ',.-' for c in proposed)))
                    cmd=self.register(pattern)
                    widget.entry.configure(validate='key',validatecommand=(cmd,'%P'))
            widget.grid(row=r,column=1,sticky='w',pady=4)
        ttk.Label(detail_body,text='Cód. RCA fica na aba principal. Produto/família, datas, percentual, quantidades e nome são calculados automaticamente.',wraplength=600).grid(row=len(extra_indices)+1,column=0,columnspan=2,sticky='w',pady=14)
        detail_body.columnconfigure(1,weight=1)
        typeframe = ttk.Frame(basic); typeframe.pack(fill='x', pady=(0,3))
        ttk.Label(typeframe, text='Aplicar para:').pack(side='left', padx=(0,12))
        ttk.Radiobutton(typeframe, text='Produto individual (coluna A)', variable=self.vars['tipo'], value='produto').pack(side='left', padx=5)
        ttk.Radiobutton(typeframe, text='Família (coluna U)', variable=self.vars['tipo'], value='familia').pack(side='left', padx=5)
        # Layout compacto: linhas fixas para evitar que a pesquisa desloque outros campos.
        fields = ttk.Frame(basic)
        fields.pack(anchor='w', pady=(0, 5))
        self._fields = fields
        fields.columnconfigure(0, minsize=352)
        fields.columnconfigure(1, minsize=352)

        def cell(label, r, c, span=1):
            f = ttk.Frame(fields)
            f.grid(row=r, column=c, columnspan=span, sticky='nw', padx=(0, 12), pady=(2, 3))
            ttk.Label(f, text=label).pack(anchor='w')
            return f

        # Linha 1: pesquisa de produto e nome da campanha.
        f = cell('Pesquisar produto', 0, 0)
        line = ttk.Frame(f); line.pack(anchor='w')
        self.product_filter = tk.StringVar(value='Código')
        self.product_filter_combo = ttk.Combobox(
            line, textvariable=self.product_filter, state='readonly', width=14,
            values=['Código', 'Descrição', 'Cód. auxiliar', 'Cód. auxiliar 2'])
        self.product_filter_combo.pack(side='left', padx=(0, 5))
        self.product_query = tk.StringVar()
        self.product_search = RoundedEntry(line, self.product_query, width=158)
        self.product_search.pack(side='left')
        self.product_search.entry.bind('<Return>', lambda e: self.search_product())
        RoundedButton(line, '⌕', command=self.search_product, width=34).pack(side='left', padx=3)
        self.product_results = ttk.Frame(f)
        self.product_selected_label = ttk.Label(f, text='Nenhum produto selecionado', foreground='#aab2c0')
        self.product_selected_label.pack(anchor='w', pady=(3, 0))
        self.product_filter_combo.bind('<<ComboboxSelected>>', lambda e: (self.product_query.set(''), self.hide_results(self.product_results), self.product_search.entry.focus_set()))

        f = cell('Nome da campanha', 0, 1)
        RoundedEntry(f, self.vars['nome'], width=215).pack(anchor='w')

        # Linha compacta: tabela, fator e preco juntos, sem espacamento entre colunas.
        price_row = ttk.Frame(fields)
        price_row.grid(row=2, column=0, columnspan=2, sticky='w', padx=(0, 10), pady=(2, 3))
        region_cell = ttk.Frame(price_row)
        region_cell.pack(side='left', anchor='n', padx=(0, 9))
        ttk.Label(region_cell, text='Tabela de preço base (região ativa)').pack(anchor='w')
        self.region_combo = ttk.Combobox(region_cell, style='TCombobox', textvariable=self.regiao,
                                         state='readonly', width=22)
        self.region_combo.pack(anchor='w', pady=3)
        self.region_combo.bind('<<ComboboxSelected>>', lambda e: self.load_price())

        f = cell('Pesquisar RCA (vazio por padrão)', 1, 1)
        line = ttk.Frame(f); line.pack(anchor='w')
        self.rca_filter = tk.StringVar(value='Código')
        self.rca_filter_combo = ttk.Combobox(line, textvariable=self.rca_filter,
                                              state='readonly', width=10, values=['Código', 'Nome'])
        self.rca_filter_combo.pack(side='left', padx=(0, 5))
        self.rca_query = tk.StringVar()
        self.rca_search = RoundedEntry(line, self.rca_query, width=145)
        self.rca_search.pack(side='left')
        self.rca_search.entry.bind('<FocusIn>', lambda e: self.start_rca_search())
        self.rca_search.entry.bind('<Return>', lambda e: self.search_rca())
        self.rca_button = RoundedButton(line, '⌕', command=self.search_rca, width=32)
        self.rca_button.pack(side='left', padx=3)
        rca_info_row = ttk.Frame(f)
        rca_info_row.pack(anchor='w', fill='x', pady=(3, 0))
        ttk.Checkbutton(rca_info_row, style='Dark.TCheckbutton', text='Todos', variable=self.todos_rca,
                        command=self.toggle_all_rca).pack(side='left')
        self.rca_results = ttk.Frame(f)
        self.rca_selected_label = ttk.Label(rca_info_row, text='RCA: vazio (todos)', foreground='#aab2c0')
        self.rca_selected_label.pack(side='left', padx=(12, 0))
        self.rca_filter_combo.bind('<<ComboboxSelected>>', lambda e: (self.rca_query.set(''), self.hide_results(self.rca_results), self.rca_search.entry.focus_set()))

        # Datas acima da tabela de preco, com inicio e fim lado a lado.
        dates = ttk.Frame(fields)
        dates.grid(row=1, column=0, sticky='nw', padx=(0, 10), pady=(1, 5))
        self.date_widgets = {}
        start = ttk.Frame(dates); start.pack(side='left', anchor='n', padx=(0, 9))
        ttk.Label(start, text='Início (DD/MM/AAAA)').pack(anchor='w')
        line = ttk.Frame(start); line.pack(anchor='w')
        self.date_widgets['inicio'] = RoundedEntry(line, self.vars['inicio'], width=116)
        self.date_widgets['inicio'].pack(side='left')
        RoundedButton(line, '📅', width=30, command=lambda: self.open_calendar('inicio')).pack(side='left', padx=2)
        end = ttk.Frame(dates); end.pack(side='left', anchor='n')
        ttk.Label(end, text='Fim (DD/MM/AAAA)').pack(anchor='w')
        line = ttk.Frame(end); line.pack(anchor='w')
        self.date_widgets['fim'] = RoundedEntry(line, self.vars['fim'], width=116)
        self.date_widgets['fim'].pack(side='left')
        self.end_calendar = RoundedButton(line, '📅', width=30, command=lambda: self.open_calendar('fim'))
        self.end_calendar.pack(side='left', padx=2)
        ttk.Checkbutton(end, style='Dark.TCheckbutton', text='Campanha infinita',
                        variable=self.infinita, command=self.toggle_infinite).pack(anchor='w', pady=(2, 0))

        unit_cell = ttk.Frame(price_row); unit_cell.pack(side='left', anchor='n', padx=(0, 9))
        ttk.Label(unit_cell, text='Unidade (QTUNIT)').pack(anchor='w')
        self.unit_combo = ttk.Combobox(unit_cell, style='TCombobox', textvariable=self.vars['unidade'],
                                       state='readonly', width=6)
        self.unit_combo.pack(anchor='w', pady=3)
        self.unit_combo.bind('<<ComboboxSelected>>', lambda e: self.update_price())
        price_cell = ttk.Frame(price_row); price_cell.pack(side='left', anchor='n')
        ttk.Label(price_cell, text='Preço de tabela convertido (R$)').pack(anchor='w')
        self.price_entry = RoundedEntry(price_cell, self.vars['tabela'], width=165)
        self.price_entry.pack(anchor='w', pady=3)
        # R$ é apenas visual: não entra no valor, no TXT ou nas consultas.
        self.price_entry.entry.place(x=35, y=6, width=119, height=19)
        tk.Label(self.price_entry, text='R$', bg='#242832', fg='#f0f2f6',
                 font=('Segoe UI', 10)).place(x=10, y=6, height=19)
        self.vars['tabela'].trace_add('write', lambda *_: self.update_discounts())

        self.rca_query.trace_add('write', self._rca_query_changed)
        self.product_query.trace_add('write', self._product_query_changed)
        self.after(400,self.load_regions)
        ttk.Label(basic,text='Faixas de preço (quantidades na unidade de venda)',font=('Segoe UI',11,'bold')).pack(anchor='w',pady=(5,3))
        head=ttk.Frame(basic);head.pack(anchor='w')
        for i,label in enumerate(['Qtde inicial','Qtde final','Preço novo (R$)','Desconto']):
            cell=ttk.Frame(head,width=(140 if i<3 else 90),height=21)
            cell.grid(row=0,column=i,sticky='w')
            cell.grid_propagate(False)
            ttk.Label(cell,text=label).place(x=4,y=2)
        self.rows_frame=ttk.Frame(basic);self.rows_frame.pack(anchor='w')
        for a,b,p in [('1','10',''),('11','20','')]:self.add_range(a,b,p)
        RoundedButton(basic,text='+ Adicionar faixa',command=self.add_range,width=135).pack(anchor='w',pady=(4,6))

    def _product_query_changed(self, *_):
        self._product_search_token += 1
        self.hide_results(self.product_results)

    def _rca_query_changed(self, *_):
        self._rca_search_token += 1
        self.hide_results(self.rca_results)

    def _db_query(self, sql, binds, callback):
        ctx=self.winthor
        if not ctx or not all(ctx.get(k) for k in ('usuario_bd','senha_bd','alias')):
            messagebox.showwarning('Oracle','Abra o Gerador pela rotina 529 do WinThor para consultar o banco.')
            return
        def worker():
            try:
                with self._oracle_lock:
                    if not oracledb:
                        raise RuntimeError('Driver Oracle indisponível')
                    tns=find_tns_directory(ctx['alias'])
                    client=bundled_oracle_client()
                    if not tns or not client:
                        raise RuntimeError('TNS ou OCI não localizado')
                    try: oracledb.init_oracle_client(lib_dir=str(client),config_dir=tns)
                    except Exception as ex:
                        if 'already been initialized' not in str(ex).lower() and 'DPY-2017' not in str(ex): raise
                    with oracledb.connect(user=ctx['usuario_bd'],password=ctx['senha_bd'],dsn=ctx['alias']) as conn:
                        with conn.cursor() as cur:
                            cur.execute(sql,binds)
                            result=cur.fetchall()
                self.after(0,lambda r=result:callback(r))
            except Exception as ex:
                match=re.search(r'(ORA-\d+|DPY-\d+|DPI-\d+)',str(ex))
                code=match.group(1) if match else type(ex).__name__
                self.after(0,lambda c=code:messagebox.showerror('Consulta Oracle',f'Falha na consulta Oracle ({c}). Verifique o filtro e as permissões.'))
        threading.Thread(target=worker,daemon=True).start()

    def hide_results(self, frame):
        # Resultados em janela flutuante, sem alterar o grid nem ser cortados pelo Canvas.
        popup = getattr(self, '_results_popup', None)
        if popup is not None:
            try:
                popup.destroy()
            except tk.TclError:
                pass
            self._results_popup = None

    def show_results(self, frame, records, select):
        self.hide_results(frame)
        anchor = self.product_search if frame is self.product_results else self.rca_search
        anchor.update_idletasks()
        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.transient(self)
        popup.configure(bg='#454b59')
        self._results_popup = popup
        width = 460
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height() + 3
        x = min(x, max(0, popup.winfo_screenwidth() - width - 10))
        height = 35 if not records else min(8, len(records)) * 24 + 25
        popup.geometry(f'{width}x{height}+{x}+{y}')
        area = tk.Frame(popup, bg='#242832', padx=2, pady=2)
        area.pack(fill='both', expand=True)
        if not records:
            tk.Label(area, text='Nenhum resultado encontrado.', bg='#242832', fg='#f0f2f6',
                     font=('Segoe UI', 10)).pack(anchor='w', padx=6, pady=4)
            return
        listbox = tk.Listbox(area, bg='#242832', fg='#f0f2f6',
                             selectbackground='#365a91', relief='flat',
                             font=('Segoe UI', 10), exportselection=False,
                             activestyle='none', highlightthickness=0)
        listbox.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(area, orient='vertical', command=listbox.yview)
        scrollbar.pack(side='right', fill='y')
        listbox.configure(yscrollcommand=scrollbar.set)
        for code, name in records:
            listbox.insert('end', f'{code} — {name}')
        def choose(event=None):
            selection = listbox.curselection()
            if selection:
                select(records[selection[0]])
        listbox.bind('<Double-Button-1>', choose)
        listbox.bind('<Return>', choose)
        listbox.bind('<MouseWheel>', lambda e: (listbox.yview_scroll(-1 if e.delta > 0 else 1, 'units'), 'break')[1])
        listbox.bind('<Escape>', lambda e: self.hide_results(frame))
        listbox.selection_set(0)
        popup.lift()
        listbox.focus_set()

    def search_product_by_code(self):
        self.product_filter.set('Código')
        self.product_query.set(self.vars['codigo'].get().strip())
        self.search_product()

    def search_product(self):
        term = self.product_query.get().strip()
        self.hide_results(self.product_results)
        if not term:
            return
        mode = self.product_filter.get()
        if mode == 'Código':
            if not term.isdigit():
                messagebox.showwarning('Produto', 'Informe um código numérico.'); return
            condition = 'P.CODPROD = :term'; bind = {'term': int(term)}
        elif mode == 'Descrição':
            condition = "INSTR(UPPER(P.DESCRICAO), :term) > 0"
            bind = {'term': term.upper()}
        elif mode == 'Cód. auxiliar':
            condition = 'TO_CHAR(P.CODAUXILIAR) = :term'; bind = {'term': term}
        else:
            condition = 'TO_CHAR(P.CODAUXILIAR2) = :term'; bind = {'term': term}
        sql = f"""SELECT P.CODPROD, P.DESCRICAO FROM PCPRODUT P
                  WHERE {condition}
                  ORDER BY CASE WHEN P.CODMARCA IS NOT NULL THEN 0 ELSE 1 END,
                           P.DESCRICAO, P.CODPROD"""
        self._product_search_token += 1
        token = self._product_search_token
        def received(rows):
            if token != self._product_search_token:
                return
            records = [(str(c), str(n)) for c, n in rows]
            if mode == 'Código' and len(records) == 1:
                self.select_product(records[0])
            else:
                self.show_results(self.product_results, records, self.select_product)
        self._db_query(sql, bind, received)

    def select_product(self,record):
        code,name=record
        self.vars['codigo'].set(code)
        self._selected_product=code
        self.hide_results(self.product_results)
        self.product_query.set(code if self.product_filter.get()=='Código' else name)
        self.product_selected_label.configure(text=f'Produto: {code} — {name}')
        self.vars['unidade'].set('')
        self.unit_combo.configure(values=[])
        self._price_raw=None
        self.vars['tabela'].set('')
        self.load_units()
        self.load_price()

    def load_units(self):
        code=self.vars['codigo'].get().strip()
        if not code.isdigit():return
        self._db_query('SELECT DISTINCT QTUNIT FROM PCEMBALAGEM WHERE CODPROD=:code AND QTUNIT>0 ORDER BY QTUNIT',{'code':int(code)},lambda rows,c=code:self._units_result(rows) if self.vars['codigo'].get()==c else None)

    def _units_result(self,rows):
        values=list(dict.fromkeys(format(Decimal(str(x[0])), 'f').rstrip('0').rstrip('.') if '.' in format(Decimal(str(x[0])), 'f') else format(Decimal(str(x[0])), 'f') for x in rows))
        self.unit_combo.configure(values=values)
        if len(values)==1:self.vars['unidade'].set(values[0])
        self.update_price()

    def load_regions(self):
        self._db_query("SELECT NUMREGIAO, REGIAO FROM PCREGIAO WHERE STATUS='A' ORDER BY REGIAO, NUMREGIAO",{},self._regions_result)

    def _regions_result(self,rows):
        current=self.regiao.get()
        self.regioes={f'{str(name).strip()} — {code}':int(code) for code,name in rows}
        values=list(self.regioes)
        self.region_combo.configure(values=values)
        if current in self.regioes:self.regiao.set(current)
        elif values:self.regiao.set(values[0])
        self.load_price()

    def load_price(self):
        code=self.vars['codigo'].get().strip()
        region=self.regioes.get(self.regiao.get())
        self._price_raw=None
        self.vars['tabela'].set('')
        if not code.isdigit() or region is None:return
        self._load_token+=1
        token=self._load_token
        self._db_query('SELECT PTABELA FROM PCTABPR WHERE CODPROD=:code AND NUMREGIAO=:region',{'code':int(code),'region':region},lambda rows:self._price_result(rows,token))

    def _price_result(self,rows,token):
        if token!=self._load_token:return
        if not rows or rows[0][0] is None:
            self._price_raw=None
            self.vars['tabela'].set('')
            return
        self._price_raw=Decimal(str(rows[0][0]))
        self.update_price()

    def update_price(self):
        if self._price_raw is None:return
        try:unit=dec(self.vars['unidade'].get())
        except Exception:
            self.vars['tabela'].set('');return
        price=self._price_raw*unit
        self.vars['tabela'].set(f'{price:.2f}'.replace('.',','))

    def update_discounts(self):
        if not hasattr(self,'ranges'):return
        try:base=dec(self.vars['tabela'].get())
        except Exception:base=Decimal('0')
        for frame,vals in self.ranges:
            if not hasattr(frame,'pct'):continue
            try:
                price=dec(vals[2].get())
                frame.pct.set(f'{((base-price)/base*100):.2f}%'.replace('.',',') if base>0 else '—')
            except Exception:frame.pct.set('—')

    def toggle_infinite(self):
        today=datetime.today().strftime('%d/%m/%Y')
        if self.infinita.get():
            self.vars['inicio'].set(today)
            self.vars['fim'].set('01/01/2050')
        else:
            self.vars['fim'].set(today)
        self.date_widgets['fim'].entry.configure(state='disabled' if self.infinita.get() else 'normal')
        self.end_calendar.configure(state='disabled' if self.infinita.get() else 'normal')

    def toggle_all_rca(self):
        state='disabled' if self.todos_rca.get() else 'normal'
        self.rca_filter_combo.configure(state='disabled' if self.todos_rca.get() else 'readonly')
        self.rca_button.configure(state=state)
        self.rca_search.entry.configure(state=state)
        self.hide_results(self.rca_results)
        self.rca_selected_label.configure(text='Todos os RCAs' if self.todos_rca.get() else f'RCA: {self.vars["rca"].get() or "vazio (todos)"}')

    def start_rca_search(self):
        if self.todos_rca.get():return
        pass  # RCA permanece vazio até seleção explícita

    def search_rca(self):
        if self.todos_rca.get(): return
        term = self.rca_query.get().strip()
        self.hide_results(self.rca_results)
        if not term:
            self.rca_selected_label.configure(text=f'RCA: {self.vars["rca"].get() or "vazio (todos)"}')
            return
        mode = self.rca_filter.get()
        if mode == 'Código':
            if not term.isdigit():
                messagebox.showwarning('RCA', 'Selecione Nome para pesquisar por nome.'); return
            condition = 'U.CODUSUR = :term'; bind = {'term': int(term)}
        else:
            condition = "INSTR(UPPER(U.NOME), :term) > 0"
            bind = {'term': term.upper()}
        sql = f"""SELECT U.CODUSUR, U.NOME FROM PCUSUARI U
                  WHERE U.DTTERMINO IS NULL AND {condition}
                  ORDER BY U.NOME, U.CODUSUR"""
        self._rca_search_token += 1
        token = self._rca_search_token
        def received(rows):
            if token != self._rca_search_token:
                return
            records = [(str(c), str(n)) for c, n in rows]
            if mode == 'Código' and len(records) == 1:
                self.select_rca(records[0])
            else:
                self.show_results(self.rca_results, records, self.select_rca)
        self._db_query(sql, bind, received)

    def select_rca(self,record):
        code,name=record
        self.vars['rca'].set(code)
        self.rca_query.set(code if self.rca_filter.get()=='Código' else name)
        self.rca_selected_label.configure(text=f'RCA: {code} — {name}')
        self.hide_results(self.rca_results)

    def test_oracle(self):
        """Consulta de leitura usando credenciais entregues pelo launcher WinThor."""
        ctx = self.winthor
        if not ctx or not all(ctx.get(k) for k in ('usuario_bd','senha_bd','alias')):
            messagebox.showwarning('Conexão Oracle', 'Abra a rotina pelo WinThor com os cinco parâmetros configurados.')
            return
        self.db_button.configure(state='disabled')
        self.db_status.set('Oracle: conectando...')
        def worker():
            try:
                if oracledb is None:
                    result = ('missing', ORACLE_IMPORT_ERROR[1] or ORACLE_IMPORT_ERROR[0])
                    self.after(0, lambda result=result: self._oracle_result(result))
                    return
                tns_dir = find_tns_directory(ctx['alias'])
                if not tns_dir:
                    result = ('tns_missing', ctx['alias'])
                    self.after(0, lambda result=result: self._oracle_result(result))
                    return
                # OCI empacotado: não usa o Oracle Client 10g/11g do Windows.
                client_dir = bundled_oracle_client()
                if client_dir is None:
                    result = ('oci_missing',)
                    self.after(0, lambda result=result: self._oracle_result(result))
                    return
                try: oracledb.init_oracle_client(lib_dir=str(client_dir), config_dir=tns_dir)
                except Exception as ex:
                    if 'already been initialized' not in str(ex).lower() and 'DPY-2017' not in str(ex): raise
                with oracledb.connect(user=ctx['usuario_bd'], password=ctx['senha_bd'],
                                      dsn=ctx['alias']) as conn:
                    with conn.cursor() as cur:
                        cur.execute('SELECT USER, SYS_CONTEXT(\'USERENV\', \'DB_NAME\') FROM DUAL')
                        db_user, db_name = cur.fetchone()
                result = ('ok', str(db_user), str(db_name))
            except ModuleNotFoundError as exc:
                result = ('missing', getattr(exc, 'name', None) or 'desconhecido')
            except Exception as exc:
                # Nunca apresentar argumentos, senha ou string de conexão no erro.
                # Código Oracle/DPY é útil sem expor detalhes potencialmente sensíveis.
                import re
                match = re.search(r'(ORA-\d+|DPY-\d+|DPI-\d+)', str(exc))
                result = ('error', match.group(1) if match else type(exc).__name__)
            self.after(0, lambda: self._oracle_result(result))
        threading.Thread(target=worker, daemon=True).start()

    def _oracle_result(self, result):
        self.db_button.configure(state='normal')
        if result[0] == 'ok':
            self.db_status.set(f'Oracle: conectado • Usuário BD: {result[1]} • Banco: {result[2]}')
            messagebox.showinfo('Conexão Oracle', 'Conexão estabelecida com sucesso!\n'
                                f'Usuário Oracle: {result[1]}\nBanco: {result[2]}\n'
                                'Teste somente leitura; nenhuma alteração foi realizada.')
        elif result[0] == 'oci_missing':
            self.db_status.set('Oracle: bibliotecas OCI não empacotadas')
            messagebox.showerror('Conexão Oracle',
                                 'Este EXE não contém a pasta oracle_client com oci.dll.\n'
                                 'Recompile incluindo o Instant Client compatível.\n'
                                 'O cliente Oracle 10g instalado no Windows não será modificado.')
        elif result[0] == 'tns_missing':
            self.db_status.set('Oracle: alias TNS não localizado')
            messagebox.showerror('Conexão Oracle',
                                 f'Não encontrei um tnsnames.ora com o alias {result[1]}.\n'
                                 'Configure TNS_ADMIN ou verifique a instalação do Oracle Client.')
        elif result[0] == 'missing':
            self.db_status.set('Oracle: driver não instalado')
            messagebox.showerror('Conexão Oracle',
                                 f'Falha ao carregar módulo: {result[1]}\n'
                                 'Verifique a inclusão do driver e suas dependências no executável.')
        else:
            self.db_status.set(f'Oracle: falha ({result[1]})')
            messagebox.showerror('Conexão Oracle', f'Não foi possível conectar ({result[1]}).\n'
                                 'Verifique credenciais, resolução do alias TNS e acesso à rede.\n'
                                 'Se necessário, configure TNS_ADMIN apontando para a pasta do tnsnames.ora.')

    def open_calendar(self, key):
        """Calendário nativo em Tkinter; data continua editável por digitação."""
        popup=tk.Toplevel(self)
        popup.title('Selecionar data')
        popup.update_idletasks()
        anchor=self.date_widgets[key]
        x=anchor.winfo_rootx();y=anchor.winfo_rooty()+anchor.winfo_height()+2
        popup.geometry(f'+{x}+{y}')
        popup.configure(bg='#101114')
        popup.resizable(False,False)
        popup.transient(self)
        try: chosen=date(self.vars[key].get())
        except ValueError: chosen=datetime.today()
        year=tk.IntVar(value=chosen.year)
        month=tk.IntVar(value=chosen.month)
        header=ttk.Frame(popup,padding=10);header.pack(fill='x')
        grid=ttk.Frame(popup,padding=(10,0,10,10));grid.pack()
        def shift(delta):
            value=year.get()*12+month.get()-1+delta
            year.set(value//12)
            month.set(value%12+1)
            draw()
        RoundedButton(header,text='◀',width=36,command=lambda:shift(-1)).pack(side='left')
        heading=ttk.Label(header,text='',width=23,anchor='center',font=('Segoe UI',10,'bold'))
        heading.pack(side='left',padx=6)
        RoundedButton(header,text='▶',width=36,command=lambda:shift(1)).pack(side='left')
        def choose(day):
            self.vars[key].set(f'{day:02d}/{month.get():02d}/{year.get():04d}')
            popup.destroy()
        def draw():
            heading.configure(text=f'{["","Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"][month.get()]} de {year.get()}')
            for widget in grid.winfo_children():widget.destroy()
            for col,day in enumerate(['Seg','Ter','Qua','Qui','Sex','Sáb','Dom']):
                ttk.Label(grid,text=day,width=5,anchor='center').grid(row=0,column=col,pady=(0,5))
            for row,week in enumerate(calendar.monthcalendar(year.get(),month.get()),1):
                for col,day in enumerate(week):
                    if day:
                        RoundedButton(grid,text=str(day),width=36,command=lambda d=day:choose(d)).grid(row=row,column=col,padx=1,pady=1)
        draw()
        popup.grab_set()
        popup.focus_set()

    def add_range(self,a='',b='',p='',as_placeholders=False):
        vals=[tk.StringVar(value=('0' if as_placeholders and i==2 else '' if as_placeholders else x)) for i,x in enumerate((a,b,p))]
        frame=ttk.Frame(self.rows_frame);frame.pack(fill='x',pady=2)
        for i,var in enumerate(vals):
            (PlaceholderEntry(frame,textvariable=var,placeholder=(a,b,'')[i],width=132) if as_placeholders and i < 2 else RoundedEntry(frame,textvariable=var,width=132)).grid(row=0,column=i,sticky='w',padx=4)
        # Prefixo visual dentro do campo, sem alterar a StringVar.
        price_widget=frame.grid_slaves(row=0,column=2)[0]
        ttk.Label(price_widget,text='R$',background='#242832',foreground='#f0f2f6').place(x=7,y=6)
        price_widget.entry.place(x=32,y=6,width=price_widget.winfo_reqwidth()-40,height=19)
        pct=tk.StringVar(value='—')
        ttk.Label(frame,textvariable=pct,width=11,anchor='w').grid(row=0,column=4,padx=5)
        frame.pct=pct
        vals[2].trace_add('write',lambda *_:self.update_discounts())
        remove_button=RoundedButton(frame,text='Remover',command=lambda:self.remove_range(frame),width=86)
        remove_button.grid(row=0,column=3,padx=4)
        frame.remove_button=remove_button
        self.ranges.append((frame,vals))
        self.update_remove_buttons()
        self.update_discounts()
    def update_remove_buttons(self):
        for frame,_ in self.ranges:
            frame.remove_button.configure(state='disabled' if len(self.ranges)==1 else 'normal')
    def remove_range(self,frame):
        if len(self.ranges)<=1:
            messagebox.showwarning('Faixa obrigatória','A política precisa ter pelo menos uma faixa de desconto.')
            return
        self.ranges=[x for x in self.ranges if x[0]!=frame]
        frame.destroy()
        self.update_remove_buttons()
        self.update_discounts()
    def collect(self):
        return make_rows({**{k:v.get() for k,v in self.vars.items()},'todos_rca':self.todos_rca.get()},[[v.get() for v in vals] for _,vals in self.ranges],{i:v.get() for i,v in self.extra_vars.items()})
    def preview(self):
        try:
            rows=self.collect()
            self.output.delete('1.0','end')
            self.output.insert('end',f'{len(rows)} linha(s) gerada(s). Colunas não preenchidas permanecem vazias.\n\n')
            for row in rows:self.output.insert('end','\t'.join(row)+'\n')
        except (ValueError,InvalidOperation) as e:messagebox.showerror('Verifique os dados',str(e))
    def save(self):
        try: rows=self.collect()
        except (ValueError,InvalidOperation) as e:
            messagebox.showerror('Verifique os dados',str(e));return
        path=filedialog.asksaveasfilename(title='Salvar TXT da rotina 561',defaultextension='.txt',filetypes=[('Arquivo TXT','*.txt')],initialfile='politica_561_teste.txt')
        if not path:return
        try:
            with open(path,'w',encoding='cp1252',newline='') as f:
                writer=csv.writer(f,delimiter='\t',lineterminator='\r\n',quoting=csv.QUOTE_MINIMAL)
                writer.writerows(rows)
            self.preview()
            messagebox.showinfo('TXT gerado',f'{len(rows)} linha(s) salvas em:\n{path}\n\nRCA: {self.vars["rca"].get() or "vazio"}. Confira antes da importação.')
        except (OSError,UnicodeEncodeError) as e:messagebox.showerror('Erro ao salvar',str(e))

if __name__=='__main__': App().mainloop()
