import os
import sys
import locale
import datetime
import pyodbc
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk


class AccessFinanceDB:
    def __init__(self, db_path=None):
        current_dir = os.path.dirname(__file__)
        default_path = os.path.join(current_dir, 'finances.mdb')
        fallback_path = os.path.join('c:\\PythonWork\\finance', 'finances.mdb')
        self.db_path = db_path or default_path
        if not os.path.exists(self.db_path) and os.path.exists(fallback_path):
            self.db_path = fallback_path
        if not os.path.exists(self.db_path):
            raise FileNotFoundError(f"Access database not found: {self.db_path}")
        self.connection = None
        self.cursor = None
        self.connect()

    def connect(self):
        dsn = rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path};'
        self.connection = pyodbc.connect(dsn)
        self.cursor = self.connection.cursor()

    def commit(self):
        if self.connection:
            self.connection.commit()

    def close(self):
        if self.cursor:
            self.cursor.close()
        if self.connection:
            self.connection.close()

    def fetch_list(self, table_name):
        sql = f"SELECT Description FROM {table_name} ORDER BY Description"
        self.cursor.execute(sql)
        return [row[0] for row in self.cursor.fetchall() if row[0] is not None]

    def get_id(self, table_name, description):
        sql = f"SELECT AutoId FROM {table_name} WHERE Description = ?"
        self.cursor.execute(sql, description)
        row = self.cursor.fetchone()
        return row[0] if row else None

    def insert_metadata(self, table_name, description):
        sql = f"INSERT INTO {table_name} (Description) VALUES (?)"
        self.cursor.execute(sql, description)
        self.commit()
        return True

    def delete_metadata(self, table_name, description):
        meta_id = self.get_id(table_name, description)
        if meta_id is None:
            return False
        fk_column = {
            'ExpanceCategory': 'AutoCategoryId',
            'expanceplace': 'AutoPlaceId',
            'PaymentTypes': 'AutoPaymentId'
        }.get(table_name)
        if fk_column:
            delete_expenses_sql = f"DELETE FROM expances WHERE {fk_column} = ?"
            self.cursor.execute(delete_expenses_sql, meta_id)
        delete_meta_sql = f"DELETE FROM {table_name} WHERE AutoId = ?"
        self.cursor.execute(delete_meta_sql, meta_id)
        self.commit()
        return True

    def insert_expense(self, amount, exp_date, category, place, payment, comment, paid):
        category_id = self.get_id('ExpanceCategory', category)
        place_id = self.get_id('expanceplace', place)
        payment_id = self.get_id('PaymentTypes', payment)
        if not all([category_id, place_id, payment_id]):
            raise ValueError('Category, Place, or Payment does not exist.')
        sql = ('INSERT INTO expances (Amount, ExpanceDate, AutoCategoryId, AutoPlaceId, '
               'AutoPaymentId, Comment, Paid) VALUES (?, ?, ?, ?, ?, ?, ?)')
        self.cursor.execute(sql, amount, exp_date, category_id, place_id, payment_id, comment, paid)
        self.cursor.execute('SELECT @@IDENTITY AS ID')
        row = self.cursor.fetchone()
        self.commit()
        return int(row[0]) if row else None

    def update_expense(self, auto_id, amount, exp_date, category, place, payment, comment, paid):
        category_id = self.get_id('ExpanceCategory', category)
        place_id = self.get_id('expanceplace', place)
        payment_id = self.get_id('PaymentTypes', payment)
        if not all([category_id, place_id, payment_id]):
            raise ValueError('Category, Place, or Payment does not exist.')
        sql = ('UPDATE expances SET Amount = ?, ExpanceDate = ?, AutoCategoryId = ?, '
               'AutoPlaceId = ?, AutoPaymentId = ?, Comment = ?, Paid = ? WHERE AutoId = ?')
        self.cursor.execute(sql, amount, exp_date, category_id, place_id, payment_id, comment, paid, auto_id)
        self.commit()
        return True

    def delete_expense(self, auto_id):
        sql = 'DELETE FROM expances WHERE AutoId = ?'
        self.cursor.execute(sql, auto_id)
        self.commit()
        return True

    def fetch_recent_expenses(self, limit=100):
        limit = int(limit or 100)
        sql = ('SELECT TOP ' + str(limit) + ' exp.AutoId, exp.Amount, exp.ExpanceDate, cat.Description AS Category, '
               'place.Description AS Place, payment.Description AS Payment, exp.Comment, exp.Paid '
               'FROM ((expances AS exp LEFT JOIN ExpanceCategory AS cat ON exp.AutoCategoryId = cat.AutoId) '
               'LEFT JOIN expanceplace AS place ON exp.AutoPlaceId = place.AutoId) '
               'LEFT JOIN PaymentTypes AS payment ON exp.AutoPaymentId = payment.AutoId '
               'ORDER BY exp.AutoId DESC')
        self.cursor.execute(sql)
        return [tuple(row) for row in self.cursor.fetchall()]


class FinanceApp(tk.Tk):
    def __init__(self, db):
        super().__init__()
        self.title('TK Finance Manager')
        self.geometry('1240x780')
        self.db = db
        self.current_selection = None
        self.create_widgets()
        self.refresh_metadata()
        self.load_recent_expenses()

    def create_widgets(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill='x')

        self._create_label_entry(top, 'Amount', 0)
        self._create_label_entry(top, 'Date', 1, default=datetime.date.today().strftime('%m/%d/%Y'))
        self._create_label_combobox(top, 'Category', 2)
        self._create_label_combobox(top, 'Place', 3)
        self._create_label_combobox(top, 'Payment', 4)
        self._create_label_entry(top, 'Paid', 5)
        self._create_label_entry(top, 'Comment', 6, width=62)

        buttons = ttk.Frame(top)
        buttons.grid(row=7, column=0, columnspan=4, sticky='w', pady=(8, 0))
        ttk.Button(buttons, text='Add Expense', command=self.add_expense).grid(row=0, column=0, padx=4)
        ttk.Button(buttons, text='Update Expense', command=self.update_expense).grid(row=0, column=1, padx=4)
        ttk.Button(buttons, text='Delete Expense', command=self.delete_expense).grid(row=0, column=2, padx=4)
        ttk.Button(buttons, text='Load Recent', command=self.load_recent_expenses).grid(row=0, column=3, padx=4)
        ttk.Button(buttons, text='Clear Fields', command=self.clear_fields).grid(row=0, column=4, padx=4)

        meta = ttk.LabelFrame(self, text='Manage Categories / Places / Payments', padding=10)
        meta.pack(fill='x', padx=10, pady=(6, 0))
        self._create_metadata_group(meta, 'Category', 0, 'ExpanceCategory')
        self._create_metadata_group(meta, 'Place', 1, 'expanceplace')
        self._create_metadata_group(meta, 'Payment', 2, 'PaymentTypes')

        table_frame = ttk.Frame(self)
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        columns = ('AutoId', 'Amount', 'Date', 'Category', 'Place', 'Payment', 'Comment', 'Paid')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', selectmode='browse')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor='w')
        self.tree.column('Comment', width=220)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        yscroll = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll.grid(row=1, column=0, sticky='ew')
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.status = tk.StringVar(value='Ready')
        ttk.Label(self, textvariable=self.status, anchor='w').pack(fill='x', padx=10, pady=(0, 10))

    def _create_label_entry(self, parent, label_text, row, width=34, default=''):
        label = ttk.Label(parent, text=f'{label_text}:')
        label.grid(row=row, column=0, sticky='w', pady=4)
        entry = ttk.Entry(parent, width=width)
        entry.grid(row=row, column=1, columnspan=3, sticky='w', padx=4, pady=4)
        entry.insert(0, default)
        setattr(self, f'txt_{label_text.lower()}', entry)

    def _create_label_combobox(self, parent, label_text, row):
        label = ttk.Label(parent, text=f'{label_text}:')
        label.grid(row=row, column=0, sticky='w', pady=4)
        combo = ttk.Combobox(parent, width=30)
        combo.grid(row=row, column=1, sticky='w', padx=4, pady=4)
        setattr(self, f'cb_{label_text.lower()}', combo)

    def _create_metadata_group(self, parent, label_text, column, table_name):
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=column, padx=10, sticky='nw')
        ttk.Label(frame, text=label_text).grid(row=0, column=0, sticky='w')
        combo = ttk.Combobox(frame, width=28)
        combo.grid(row=1, column=0, pady=4)
        ttk.Button(frame, text=f'Add {label_text}', command=lambda: self.add_metadata(table_name, combo)).grid(row=2, column=0, pady=2, sticky='ew')
        ttk.Button(frame, text=f'Delete {label_text}', command=lambda: self.delete_metadata(table_name, combo)).grid(row=3, column=0, pady=2, sticky='ew')
        setattr(self, f'cb_meta_{label_text.lower()}', combo)

    def set_status(self, message, error=False):
        self.status.set(message)

    def refresh_metadata(self):
        self.cb_category['values'] = self.db.fetch_list('ExpanceCategory')
        self.cb_place['values'] = self.db.fetch_list('expanceplace')
        self.cb_payment['values'] = self.db.fetch_list('PaymentTypes')
        self.cb_meta_category['values'] = self.db.fetch_list('ExpanceCategory')
        self.cb_meta_place['values'] = self.db.fetch_list('expanceplace')
        self.cb_meta_payment['values'] = self.db.fetch_list('PaymentTypes')

    def clear_fields(self):
        self.txt_amount.delete(0, tk.END)
        self.txt_date.delete(0, tk.END)
        self.txt_date.insert(0, datetime.date.today().strftime('%m/%d/%Y'))
        self.txt_paid.delete(0, tk.END)
        self.txt_comment.delete(0, tk.END)
        self.cb_category.set('')
        self.cb_place.set('')
        self.cb_payment.set('')
        self.cb_meta_category.set('')
        self.cb_meta_place.set('')
        self.cb_meta_payment.set('')
        self.current_selection = None
        self.set_status('Fields cleared')

    def add_metadata(self, table_name, entry):
        text = entry.get().strip()
        if not text:
            self.set_status('Value required', True)
            return
        self.db.insert_metadata(table_name, text)
        self.refresh_metadata()
        entry.delete(0, tk.END)
        self.set_status(f'{table_name} added')

    def delete_metadata(self, table_name, entry):
        text = entry.get().strip()
        if not text:
            self.set_status('Value required', True)
            return
        if messagebox.askokcancel('Confirm delete', f'Delete {text} and related expenses?'):
            success = self.db.delete_metadata(table_name, text)
            if success:
                self.refresh_metadata()
                self.load_recent_expenses()
                # Clear the corresponding combobox
                if table_name == 'ExpanceCategory':
                    self.cb_meta_category.set('')
                elif table_name == 'expanceplace':
                    self.cb_meta_place.set('')
                elif table_name == 'PaymentTypes':
                    self.cb_meta_payment.set('')
                self.set_status(f'{table_name} deleted')
            else:
                self.set_status('Delete failed', True)

    def add_expense(self):
        try:
            amount = self.txt_amount.get().strip()
            exp_date = self.txt_date.get().strip()
            category = self.cb_category.get().strip() or self.txt_category.get().strip()
            place = self.cb_place.get().strip() or self.txt_place.get().strip()
            payment = self.cb_payment.get().strip() or self.txt_payment.get().strip()
            comment = self.txt_comment.get().strip()
            paid = self.txt_paid.get().strip()
            if not all([amount, exp_date, category, place, payment]):
                raise ValueError('Amount, Date, Category, Place, Payment are required')
            auto_id = self.db.insert_expense(amount, exp_date, category, place, payment, comment, paid)
            self.load_recent_expenses(select_id=auto_id)
            self.set_status(f'Expense {auto_id} added')
        except Exception as exc:
            self.set_status(str(exc), True)

    def update_expense(self):
        if not self.current_selection:
            self.set_status('Select an expense row to update', True)
            return
        try:
            auto_id = int(self.current_selection)
            amount = self.txt_amount.get().strip()
            exp_date = self.txt_date.get().strip()
            category = self.cb_category.get().strip() or self.txt_category.get().strip()
            place = self.cb_place.get().strip() or self.txt_place.get().strip()
            payment = self.cb_payment.get().strip() or self.txt_payment.get().strip()
            comment = self.txt_comment.get().strip()
            paid = self.txt_paid.get().strip()
            if not all([amount, exp_date, category, place, payment]):
                raise ValueError('Amount, Date, Category, Place, Payment are required')
            self.db.update_expense(auto_id, amount, exp_date, category, place, payment, comment, paid)
            self.load_recent_expenses(select_id=auto_id)
            self.set_status(f'Expense {auto_id} updated')
        except Exception as exc:
            self.set_status(str(exc), True)

    def delete_expense(self):
        if not self.current_selection:
            self.set_status('Select an expense row to delete', True)
            return
        if messagebox.askokcancel('Confirm delete', f'Delete expense {self.current_selection}?'):
            try:
                self.db.delete_expense(int(self.current_selection))
                self.load_recent_expenses()
                self.clear_fields()
                self.set_status(f'Expense {self.current_selection} deleted')
            except Exception as exc:
                self.set_status(str(exc), True)

    def load_recent_expenses(self, limit=100, select_id=None):
        for item in self.tree.get_children():
            self.tree.delete(item)
        expenses = self.db.fetch_recent_expenses(limit)
        for record in expenses:
            values = [self._format_value(value) for value in record]
            self.tree.insert('', 'end', iid=record[0], values=values)
        self.set_status(f'Loaded {len(expenses)} recent expenses')
        if select_id and self.tree.exists(select_id):
            self.tree.selection_set(select_id)
            self.tree.see(select_id)

    def _format_value(self, value):
        if isinstance(value, float):
            try:
                return locale.currency(value, grouping=True)
            except Exception:
                return f'{value:.2f}'
        if isinstance(value, datetime.date):
            return value.strftime('%m/%d/%Y')
        return str(value) if value is not None else ''

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item_id = selected[0]
        values = self.tree.item(item_id, 'values')
        self.current_selection = item_id
        self.txt_amount.delete(0, tk.END)
        self.txt_date.delete(0, tk.END)
        self.txt_paid.delete(0, tk.END)
        self.txt_comment.delete(0, tk.END)
        self.txt_amount.insert(0, values[1])
        self.txt_date.insert(0, values[2])
        self.cb_category.set(values[3])
        self.cb_place.set(values[4])
        self.cb_payment.set(values[5])
        self.txt_comment.insert(0, values[6])
        self.txt_paid.insert(0, values[7])
        self.set_status(f'Selected expense {item_id}')


def main():
    try:
        locale.setlocale(locale.LC_ALL, '')
    except Exception:
        pass
    try:
        db = AccessFinanceDB()
    except Exception as exc:
        messagebox.showerror('Database error', str(exc))
        sys.exit(1)
    app = FinanceApp(db)
    app.protocol('WM_DELETE_WINDOW', lambda: (db.close(), app.destroy()))
    app.mainloop()


if __name__ == '__main__':
    main()
