import pandas as pd
import sys, pathlib
from PyQt6 import QtPrintSupport
from yattag import Doc
from PyQt6.QtWidgets import QWidget, QApplication, QFileDialog, QTreeWidgetItem, QStyledItemDelegate, QSpinBox, QDoubleSpinBox, QMessageBox, QVBoxLayout
from PyQt6.QtCore import Qt, pyqtSlot
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QBrush, QColor, QTextDocument
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from form_tv import Ui_Form

class Window(QWidget, Ui_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Payment match')
        self.setStyleSheet('QLineEdit {background-color: bisque}')
        self.btn_close.clicked.connect(app.exit)
        self.btn_open_data.clicked.connect(self.load_data)
        self.btn_open_initials.clicked.connect(self.load_initials)
        self.btn_calculate.clicked.connect(self.calculate_checked)
        self.tv_selected.itemSelectionChanged.connect(self.details)
        self.btn_print.clicked.connect(self.print_selected)
        self.btn_print_selected.clicked.connect(self.print_details)
        self.btn_save.clicked.connect(self.save_selected)
        self.btn_save_selected.clicked.connect(self.save_details)
        self.model = QStandardItemModel()
        self.tv_totals.setModel(self.model)
        self.tv_totals.setItemDelegateForColumn(3, SpinDelegate(self))
        self.tv_totals.setItemDelegateForColumn(4, DoubleSpinDelegate(self))
        self.data = self.initials = False
        vl = QVBoxLayout(self.plot_selected)
        vl.setContentsMargins(0,0,0,0)
        self.figure = Figure(figsize=(1.5,1.0), layout='constrained')
        self.canvas = FigureCanvas(self.figure)
        vl.addWidget(self.canvas)
        vl_2 = QVBoxLayout(self.plot_details)
        vl_2.setContentsMargins(0, 0, 0, 0)
        self.figure_details = Figure(figsize=(1.5, 1.0), layout='constrained')
        self.canvas_details = FigureCanvas(self.figure_details)
        vl_2.addWidget(self.canvas_details)
    def load_data(self):
        self.data_file = QFileDialog.getOpenFileName(self)[0]
        if not self.data_file: return
        self.le_data.setText(pathlib.Path(self.data_file).name)
        cols = ['Дата', 'Документ', 'Субконто1 Дт', 'Субконто2 Дт', 'Субконто1 Кт', 'Субконто2 Кт', 'Счет Дт', 'Счет Кт', 'Сумма', 'Содержание']
        conv = {'Счет Дт': str, 'Счет Кт': str}
        try:
            df = pd.read_excel(self.data_file, sheet_name='Лист_1', usecols=cols, parse_dates=['Дата'], date_format='%d.%m.%Y',converters=conv)
            df_filt = df.loc[df['Содержание'] != "Зачет аванса покупателя"]
            self.df_dt = (df_filt[df_filt['Счет Дт'].str.contains("62").fillna(False)]
                     .rename(columns={'Субконто1 Дт': 'Контрагенты', 'Субконто2 Дт': 'Договоры'})
                     .assign(операция='Д'))
            self.df_kt = (df_filt[df_filt['Счет Кт'].str.contains("62").fillna(False)]
                     .rename(columns={'Субконто1 Кт': 'Контрагенты', 'Субконто2 Кт': 'Договоры'})
                     .assign(операция='К'))
            start_date, end_date = self.df_dt['Дата'].agg(['min', 'max'])
        except Exception as err:
            QMessageBox.critical(self, 'Ошибка загрузки', str(err))
            return
        self.de_start.setDate(start_date)
        self.de_end.setDate(end_date)
        self.data = True

        def f1(df_):
            sum_d = df_.loc[df_['операция'] == 'Д', 'Сумма'].sum()
            sum_k = df_.loc[df_['операция'] == 'К', 'Сумма'].sum()
            return pd.Series([sum_d, sum_k], ['оборот Дт', 'оборот Кт'])
        def f2(df_):
            df_ = df_.reset_index().drop('Контрагенты', axis=1).sort_values('оборот Дт',ascending=False)
            return pd.Series([df_['оборот Дт'].sum(), df_['оборот Кт'].sum(), df_],
                             ['оборот Дт', 'оборот Кт', 'договоры'])

        df_customers = (pd.concat([self.df_dt, self.df_kt]).groupby(['Контрагенты', 'Договоры']).apply(f1, include_groups=False)
                        .groupby('Контрагенты').apply(f2, include_groups=False).sort_values('оборот Дт', ascending=False))
        for row in df_customers.itertuples():
            item = StandardItem(row.Index)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsAutoTristate)
            item.setCheckState(Qt.CheckState.Unchecked)
            item_1, item_2 = StandardItem(f'{row._1:_.2f}'), StandardItem(f'{row._2:_.2f}')
            for i in (item_1, item_2): i.setTextAlignment(Qt.AlignmentFlag.AlignRight)
            self.model.appendRow([item, item_1, item_2])
            for contract in row.договоры.itertuples():
                item_c = StandardItem(contract.Договоры)
                item_c.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsAutoTristate)
                item_c.setCheckState(Qt.CheckState.Unchecked)
                item1, item2, item3, item4 = StandardItem(f'{contract._2:_.2f}'), StandardItem(f'{contract._3:_.2f}'), StandardItem(), StandardItem()
                for i in (item1, item2, item3, item4): i.setTextAlignment(Qt.AlignmentFlag.AlignRight)
                item.appendRow([item_c, item1, item2, item3, item4])
        self.model.setColumnCount(5)
        self.model.setHorizontalHeaderLabels(['Контрагент, Договор', 'Оборот Дт', 'Оборот Кт', 'Срок', '% пеня'])
        self.tv_totals.setAlternatingRowColors(True)
        self.tv_totals.header().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tv_totals.resizeColumnToContents(0)
        if self.tv_totals.columnWidth(0) > 400: self.tv_totals.setColumnWidth(0, 400)
        for col in [3,4]: self.tv_totals.setColumnWidth(col,50)
        self.le_data.setStyleSheet('background-color: lightgreen')
        if self.data and self.initials: self.btn_calculate.setEnabled(True)
    def load_initials(self):
        self.initials_file = QFileDialog.getOpenFileName(self)[0]
        if not self.initials_file: return
        self.le_initials.setText(pathlib.Path(self.initials_file).name)
        columns = ['Субконто1', 'Субконто2', 'Субконто3', 'Субконто3.Дата', 'Счет', 'Сумма Конечный остаток Дт', 'Сумма Конечный остаток Кт']
        try:
            df_ost = (pd.read_excel(self.initials_file, usecols=columns, parse_dates=['Субконто3.Дата'], date_format='%d.%m.%Y').query('Счет.str.contains("62")')
                      .rename({'Субконто1': 'Контрагенты', 'Субконто2': 'Договоры',  'Субконто3':'Документ', 'Субконто3.Дата':'Дата'}, axis=1))
                      # .assign(Дата=start_date - pd.Timedelta(1, 'day'), Документ='входящее сальдо'))
            df_ost_dt = (df_ost.dropna(subset='Сумма Конечный остаток Дт')
                         .rename({'Сумма Конечный остаток Дт': 'Сумма', 'Счет': 'Счет Дт'}, axis=1)
                         .drop('Сумма Конечный остаток Кт', axis=1))
            df_ost_kt = (df_ost.dropna(subset='Сумма Конечный остаток Кт')
                         .rename({'Сумма Конечный остаток Кт': 'Сумма', 'Счет': 'Счет Кт'}, axis=1)
                         .drop('Сумма Конечный остаток Дт', axis=1))
        except Exception as err:
            QMessageBox(self, 'Загрузка входящих остатков', str(err))
            return
        self.dt_gr_df = pd.concat([df_ost_dt, self.df_dt], ignore_index=True).groupby(['Контрагенты', 'Договоры'])
        self.kt_gr_df = pd.concat([df_ost_kt, self.df_kt], ignore_index=True).groupby(['Контрагенты', 'Договоры'])
        self.initials = True
        self.le_initials.setStyleSheet('background-color: lightgreen')
        if self.data and self.initials: self.btn_calculate.setEnabled(True)
    def calculate_checked(self):
        if not (self.data and self.initials):
            QMessageBox.information(self,'Расчет результатов', 'Не загружены данные или входящие остатки.')
            return
        self.pay_df = pd.DataFrame()
        y, m, d = self.de_end.date().getDate()
        end_date = pd.Timestamp(y, m, d)
        checked = []
        for row in range(self.model.rowCount()):
            parent = self.model.invisibleRootItem().child(row)
            for r in range(parent.rowCount()):
                if parent.child(r).checkState() == Qt.CheckState.Checked:
                    child = parent.child(r)
                    v2 = parent.child(r,3).data(Qt.ItemDataRole.DisplayRole)
                    v3 = parent.child(r,4).data(Qt.ItemDataRole.DisplayRole)
                    checked.append([(parent.data(Qt.ItemDataRole.DisplayRole), child.data(Qt.ItemDataRole.DisplayRole)),v2,v3])
        if not checked:
            QMessageBox.information(self,'Загрузка данных', 'Данные для расчета не выбраны.')
            return
        self.checked_df = pd.DataFrame(checked, columns=['items', 'period', 'fine']).set_index('items').fillna({'period':self.sb_period.value(), 'fine': self.dsb_fine.value()})
        self.checked_df.index = pd.MultiIndex.from_tuples(self.checked_df.index)
        def f_4(s, pay_delay, fee):
            if self.pay_df.empty:
                d = (end_date - s['Дата']).days
                p = s['Сумма']
                f = 0 if d < pay_delay else round((d - pay_delay) * fee * p,2)
                return pd.DataFrame([[end_date, 'задолженность', p, d, f]], columns=['Дата', 'Документ', 'Сумма', 'срок оплаты', 'пеня'])
            sell = s['Сумма']
            cum_df = self.pay_df.assign(cums=self.pay_df['Сумма'].cumsum())
            pay = cum_df.query('cums<=@sell')
            pay = pay if not pay.empty else cum_df.head(1)
            self.pay_df.drop(pay.index, inplace=True)
            if not pay.empty and sell > pay.loc[pay.index[-1], 'cums']:
                ad_pay = self.pay_df.head(1)
                ad_pay['Сумма'] = round(sell - pay.loc[pay.index[-1], 'cums'],2)
                self.pay_df.head(1)['Сумма'] -= ad_pay['Сумма']
                pay = pd.concat([pay, ad_pay])
            elif sell < pay.loc[pay.index[-1], 'cums']:
                ad_pay = pay
                ad_pay['Сумма'] = round(pay.loc[pay.index[-1], 'cums'] - sell,2)
                self.pay_df = pd.concat([ad_pay, self.pay_df])
                pay['Сумма'] = sell
            pay['срок оплаты'] = (pay['Дата'] - s['Дата']).dt.days
            pay['пеня'] = pay.apply(lambda s: 0 if s['срок оплаты'] < pay_delay else round(s['Сумма'] * fee * (s['срок оплаты'] - pay_delay),2), axis=1)
            if round(pay['Сумма'].sum(),2) < sell:
                d = (end_date - s['Дата']).days
                p = round(sell - pay['Сумма'].sum(),2)
                f = 0 if d < pay_delay else round((d - pay_delay) * fee * p,2)
                pay.loc[-1, ['Дата', 'Документ', 'Сумма', 'срок оплаты', 'пеня']] = [end_date, 'задолженность', p, d, f]
            return pay[['Дата', 'Документ', 'Сумма', 'срок оплаты', 'пеня']]
        self.res_df = {}
        for x, s in self.checked_df.iterrows():
            v1, v2 = s.loc[['period', 'fine']]
            self.pay_df = self.kt_gr_df.get_group(x).sort_values('Дата') if x in self.kt_gr_df.groups.keys() else pd.DataFrame()
            if x not in self.dt_gr_df.groups.keys(): continue
            self.res_df[x] = (self.dt_gr_df.get_group(x).sort_values('Дата').
                assign(pay=lambda d: d.apply(f_4, axis=1, args=(v1, v2)),
                       пеня=lambda d: d.apply(lambda s: s.pay['пеня'].sum(), axis=1),
                       оплата=lambda d: d.apply(lambda s: s.pay.loc[s.pay['Документ'] != 'задолженность', 'Сумма'].sum(), axis=1)).
                assign(средн_срок_оплаты=lambda d: d.apply(lambda s: round((s.pay['срок оплаты'] * s.pay['Сумма']).sum() / s.pay['Сумма'].sum()), axis=1)).
                assign(max_срок_оплаты=lambda d: d.apply(lambda s: s.pay['срок оплаты'].max(), axis=1),
                       min_срок_оплаты=lambda d: d.apply(lambda s: s.pay['срок оплаты'].min(), axis=1))
                [['Дата', 'Документ', 'Счет Дт', 'Счет Кт', 'Сумма', 'оплата', 'min_срок_оплаты', 'средн_срок_оплаты', 'max_срок_оплаты', 'пеня', 'pay']])

        def f5(df_):
            min_ = df_['min_срок_оплаты'].min()
            max_ = df_['max_срок_оплаты'].max()
            avg_ = round((df_['Сумма'] * df_['средн_срок_оплаты']).sum() / df_['Сумма'].sum())
            sum_ = df_['Сумма'].sum()
            pay_ = df_['оплата'].sum()
            fine_ = df_['пеня'].sum()
            return pd.Series([min_, max_, avg_, sum_, pay_, sum_ - pay_, fine_], index=['min', 'max', 'avg', 'sum', 'pay', 'debt', 'fine'])
        def f6(df_):
            df_ = df_.reset_index().drop('контрагент', axis=1)
            min_ = df_['min'].min()
            max_ = df_['max'].max()
            avg_ = round((df_['sum'] * df_['avg']).sum() / df_['sum'].sum())
            sum_ = df_['sum'].sum()
            pay_ = df_['pay'].sum()
            fine_ = df_['fine'].sum()
            return pd.Series([min_, max_, avg_, sum_, pay_, sum_-pay_, fine_, df_], index=['min_', 'max_', 'avg_', 'sum_', 'pay_', 'debt_', 'fine_', 'contract'])

        if self.res_df:
            self.cust_df = (pd.concat([v.assign(контрагент=k1, договор=k2) for (k1, k2), v in self.res_df.items()])
                       .groupby(['контрагент', 'договор']).apply(f5, include_groups=False)
                       .groupby('контрагент').apply(f6, include_groups=False))
            self.plot_hist()
            self.show_result()
            self.btn_print.setEnabled(True)
            self.btn_save.setEnabled(True)
            self.btn_save_selected.setEnabled(False)
    def plot_hist(self):
        x = pd.concat([pd.concat([p['срок оплаты'] for p in r.pay.values]) for r in self.res_df.values()])
        self.figure.clear()
        ax = self.canvas.figure.subplots()
        ax.hist(x, bins=30, linewidth=0.5, edgecolor="white")
        ax.set(yticks=())
        ax.tick_params(labelsize=5)
        for pos, spine in ax.spines.items():
            if pos != 'bottom': spine.set_visible(False)
        self.canvas.draw()
    def show_result(self):
        self.tv_selected.clear()
        self.tv_selected.setColumnCount(8)
        self.tv_selected.setStyleSheet('QTreeWidget::item:has-children {background: lightgray}')
        self.tv_selected.setHeaderLabels(['Контрагент, Договор', 'Мин', 'Макс', 'Средн', 'Реализация', 'Оплата', 'Долг', 'Неустройка'])
        for row in self.cust_df.itertuples():
            item = QTreeWidgetItem(self.tv_selected, [row.Index, f'{int(row.min_)}', f'{int(row.max_)}', str(row.avg_), f'{row.sum_:_.2f}', f'{row.pay_:_.2f}', f'{row.debt_:_.2f}', f'{row.fine_:_.2f}'])
            child_items = [QTreeWidgetItem([d, str(int(v1)), str(int(v2)),str(v3),f'{v4:_.2f}', f'{v5:_.2f}',f'{v6:_.2f}',f'{v7:_.2f}']) for d, v1, v2, v3, v4, v5, v6, v7 in row.contract.values]
            for child_item in child_items:
                for col in range(1, 9): child_item.setTextAlignment(col, Qt.AlignmentFlag.AlignRight)
            item.addChildren(child_items)
            item.setExpanded(True)
            for col in range(1,9): item.setTextAlignment(col, Qt.AlignmentFlag.AlignRight)
        self.tv_selected.resizeColumnToContents(0)
        if self.tv_selected.columnWidth(0) > 400: self.tv_totals.setColumnWidth(0, 400)
        for col in [1,2,3]: self.tv_selected.setColumnWidth(col,45)
        for col in [4,5,6,7]: self.tv_selected.setColumnWidth(col, 100)
        self.tv_selected.setColumnWidth(0, self.tv_selected.width()-560)
        self.tv_selected.setAlternatingRowColors(True)
        self.tv_selected.header().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tv_details.clear()
        self.figure_details.clear()
        self.canvas_details.draw()
    def details(self):
        if parent := self.tv_selected.currentItem().parent():
            self.tv_details.clear()
            self.tv_details.setColumnCount(5)
            self.tv_details.setStyleSheet('QTreeWidget::item:has-children {background: lightgray}')
            self.tv_details.setHeaderLabels(['Документ','Дата','Сумма','Срок','Неустойка'])
            self.tv_details.header().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
            selection = parent.text(0), self.tv_selected.currentItem().text(0)
            selection_df = self.res_df.get(selection)
            if selection_df is None: return
            for row in selection_df.itertuples():
                item = QTreeWidgetItem(self.tv_details, [row.Документ, row.Дата.strftime('%d.%m.%Y'), f'{row.Сумма:_.2f}', str(int(row.средн_срок_оплаты)),f'{row.пеня:_.2f}'])
                child_items = [QTreeWidgetItem([d, v1.strftime('%d.%m.%Y'), f'{v2:_.2f}', str(v3), f'{v4:_.2f}']) for v1,d,v2,v3,v4 in row.pay.values]
                for child_item in child_items:
                    for col in range(1, 5): child_item.setTextAlignment(col, Qt.AlignmentFlag.AlignRight)
                item.addChildren(child_items)
                item.setExpanded(True)
                for col in range(1, 5): item.setTextAlignment(col, Qt.AlignmentFlag.AlignRight)
            self.tv_details.setColumnWidth(0, self.tv_details.width() - 425)
            self.plot_details_hist(selection_df)
            self.btn_print_selected.setEnabled(True)
            self.btn_save_selected.setEnabled(True)
    def plot_details_hist(self, selection_df):
        x = pd.concat(p['срок оплаты'] for p in selection_df.pay.values)
        self.figure_details.clear()
        ax = self.canvas_details.figure.subplots()
        ax.hist(x, bins=30, linewidth=0.5, edgecolor="white")
        ax.set(yticks=())
        ax.tick_params(labelsize=5)
        for pos, spine in ax.spines.items():
            if pos != 'bottom': spine.set_visible(False)
        self.canvas_details.draw()
    def print_selected(self):
        style_sheet = '''table {border-collapse:collapse}
                        th {background-color:lightblue; border: 1px solid gray; height:2em}
                        td {border: 1px solid gray; padding:0 1em 0 1em; text-align:right}
                        tr.cust {background-color:lightgray; font-weight:bold}
                        td.D {padding-left:2em; text-align:left}
                        td.C {text-align:left}'''
        text_doc = QTextDocument()
        text_doc.setDefaultStyleSheet(style_sheet)
        text_doc.setHtml(self.get_text_doc())
        prev_dialog = QtPrintSupport.QPrintPreviewDialog()
        prev_dialog.paintRequested.connect(text_doc.print)
        prev_dialog.exec()
    def get_text_doc(self):
        doc, tag, text, line = Doc().ttl()
        with tag('html'):
            with tag('table', style='float:right'):
                doc.asis(f'<tr><td>начало периода</td><td>{self.de_start.date().toString("dd.MM.yyyy")}</td></tr>'
                         f'<tr><td>конец периода</td><td>{self.de_end.date().toString("dd.MM.yyyy")}</td></tr>')
            doc.asis('<br><br><br>')
            with tag('table', klass='checked'):
                with tag('tr', klass='head'):
                    doc.asis('<th>Контрагент</th><th>Договор</th><th>Срок оплаты</th><th>Неустойка</>')
                for (i, j), r in self.checked_df.iterrows():
                    with tag('tr'):
                        line('td', i, klass='C')
                        line('td', j, klass='C')
                        line('td', r['period'])
                        line('td', r['fine'])
            doc.stag('br')
            with tag('table', klass='tbl'):
                with tag('tr', klass='head'):
                    doc.asis('<th>Контрагент, Договор</th><th>Мин</th><th>Макс</th><th>Средн</th><th>Реализация</th><th>Оплата</th><th>Долг</th><th>Пеня</th>')
                for row in self.cust_df.itertuples():
                    with tag('tr', klass='cust'):
                        line('td', row.Index, klass='C')
                        line('td', f'{int(row.min_)}')
                        line('td', f'{int(row.max_)}')
                        line('td', f'{int(row.avg_)}')
                        line('td', f'{row.sum_:_.2f}')
                        line('td', f'{row.pay_:_.2f}')
                        line('td', f'{row.debt_:_.2f}')
                        line('td', f'{row.fine_:_.2f}')
                    for _, r in row.contract.iterrows():
                        with tag('tr', klass='K'):
                            line('td', r['договор'], klass='D')
                            line('td', f'{int(r["min"])}')
                            line('td', f'{int(r["max"])}')
                            line('td', f'{int(r["avg"])}')
                            line('td', f'{r["sum"]:_.2f}')
                            line('td', f'{r["pay"]:_.2f}')
                            line('td', f'{r["debt"]:_.2f}')
                            line('td', f'{r["fine"]:_.2f}')
        return doc.getvalue()
    def print_details(self):
        style_sheet = '''table {border-collapse:collapse}
                         th {background-color:lightblue; border: 1px solid gray; height:2em}
                         td {border: 1px solid gray; padding:0 1em 0 1em; text-align:right}
                         tr.cust {background-color:lightgray; font-weight:bold}
                         td.D {padding-left:2em; text-align:left}
                         td.C {text-align:left}'''
        text_doc = QTextDocument()
        text_doc.setDefaultStyleSheet(style_sheet)
        text_doc.setHtml(self.get_details_doc())
        prev_dialog = QtPrintSupport.QPrintPreviewDialog()
        prev_dialog.paintRequested.connect(text_doc.print)
        prev_dialog.exec()
    def get_details_doc(self):
        doc, tag, text, line = Doc().ttl()
        current_item = self.tv_selected.currentItem()
        selection = self.tv_selected.currentItem().parent().text(0), self.tv_selected.currentItem().text(0)
        with tag('html'):
            with tag('table', klass='checked'):
                with tag('tr', klass='head'):
                    doc.asis('<th>Контрагент</th><th>Договор</th><th>Мин</th><th>Макс</th><th>Средн</th><th>Реализация</th><th>Оплата</th><th>Долг</th><th>Пеня</th>')
                with tag('tr'):
                    line('td', self.tv_selected.currentItem().parent().text(0), klass='C')
                    for i in range(self.tv_selected.columnCount()): line('td', current_item.text(i))
            doc.stag('br')
            with tag('table', klass='details'):
                with tag('tr', klass='head'):
                    doc.asis('<th>Документ</th><th>Дата</th><th>Сумма</th><th>Срок</th><th>Неустойка</th>')
                    for row in self.res_df.get(selection).itertuples():
                        with tag('tr', klass='cust'):
                            line('td', row.Документ, klass='C')
                            line('td', row.Дата.strftime('%d.%m.%Y'))
                            line('td', f'{row.Сумма:_.2f}')
                            line('td', str(row.средн_срок_оплаты))
                            line('td', f'{row.пеня:_.2f}')
                        for r in row.pay.itertuples():
                            with tag('tr', klass='K'):
                                line('td', r.Документ, klass='D')
                                line('td', r.Дата.strftime('%d.%m.%Y'))
                                line('td', f'{r.Сумма:_.2f}')
                                line('td', str(r._4))
                                line('td', f'{r.пеня:_.2f}')
        return doc.getvalue()
    def save_selected(self):
        file_name, _ = QFileDialog.getSaveFileName(self, 'Save as xlsx', '', 'Excel files(*.xlsx)')
        if file_name:
            self.checked_df.to_excel(file_name, sheet_name='settings')
            with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                self.cust_df.drop('contract', axis=1).to_excel(writer, sheet_name='customers')
                r = 0
                for row in self.cust_df.itertuples():
                    row.contract.assign(контрагент=row.Index).set_index(['контрагент', 'договор']).to_excel(writer, sheet_name='contracts', startrow=r + bool(r), header=(r == 0))
                    r += row.contract.shape[0]
    def save_details(self):
        file_name, _ = QFileDialog.getSaveFileName(self, 'Save as xlsx', '', 'Excel files(*.xlsx)')
        if file_name:
            selected_cust, selected_contr = self.tv_selected.currentItem().parent().text(0), self.tv_selected.currentItem().text(0)
            sel_cust_df = self.cust_df.loc[selected_cust, 'contract'].query('договор==@selected_contr').assign(контрагент=selected_cust).set_index(['контрагент', 'договор'])
            sel_cust_df.to_excel(file_name, sheet_name='contract')
            selected_df = self.res_df[(selected_cust, selected_contr)]
            with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='overlay', date_format='DD-MM-YYYY') as writer:
                selected_df.drop('pay', axis=1).to_excel(writer, sheet_name='sales')
                r = 0
                for row in selected_df.itertuples():
                    row.pay.assign(реализация=row.Документ).set_index(['реализация', 'Документ']).to_excel(writer, sheet_name='payments', startrow=r + bool(r), header=(r == 0))
                    r += row.pay.shape[0]

class SpinDelegate(QStyledItemDelegate):
    def __init__(self, parent):
        super().__init__(parent)
    def createEditor(self, parent, option, index):
        spin = QSpinBox(parent)
        spin.setMinimum(0)
        spin.setMaximum(100)
        spin.setSingleStep(1)
        is_parent = index.model().hasChildren(index.sibling(index.row(), 0))
        if not is_parent: return spin
    def setEditorData(self, editor, index):
        editor.setValue(10)
    def setModelData(self, editor, model, index):
        model.setData(index, editor.value())

    @pyqtSlot()
    def currentIndexChanged(self):
        self.commitData.emit(self.sender())

class DoubleSpinDelegate(QStyledItemDelegate):
    def __init__(self, parent):
        super().__init__(parent)
    def createEditor(self, parent, option, index):
        double_spin = QDoubleSpinBox(parent)
        double_spin.setMinimum(0)
        double_spin.setMaximum(0.1)
        double_spin.setSingleStep(0.0001)
        double_spin.setDecimals(4)
        is_parent = index.model().hasChildren(index.sibling(index.row(), 0))
        if not is_parent: return double_spin
    def setEditorData(self, editor, index):
        editor.setValue(0.001)
    def setModelData(self, editor, model, index):
        model.setData(index, editor.value())

    @pyqtSlot()
    def currentIndexChanged(self):
        self.commitData.emit(self.sender())

class StandardItem(QStandardItem):
    def data(self, role=Qt.ItemDataRole.UserRole+1):
        if role == Qt.ItemDataRole.CheckStateRole and self.hasChildren() and self.flags() & Qt.ItemFlag.ItemIsAutoTristate:
            return self._children_check_state()
        return super().data(role)
    def _children_check_state(self):
        checked = unchecked = False
        for row in range(self.rowCount()):
            child = self.child(row)
            value = child.data(Qt.ItemDataRole.CheckStateRole)
            if value is None: return
            elif value == 0: unchecked = True
            elif value == 2: checked = True
            if unchecked and checked:
                return Qt.CheckState.PartiallyChecked
        if unchecked: return Qt.CheckState.Unchecked
        elif checked: return Qt.CheckState.Checked
    def setData(self, value, role = Qt.ItemDataRole.UserRole+1):
        if role == Qt.ItemDataRole.CheckStateRole:
            if self.flags() & Qt.ItemFlag.ItemIsAutoTristate and value != Qt.CheckState.PartiallyChecked:
                for row in range(self.rowCount()):
                    child = self.child(row)
                    if child.data(role) is not None:
                        flags = self.flags()
                        self.setFlags(flags & ~Qt.ItemFlag.ItemIsAutoTristate)
                        child.setData(value, role)
                        self.setFlags(flags)
                model = self.model()
                if model is not None:
                    parent = self
                    while True:
                        parent = parent.parent()
                        if parent is not None and parent.flags() & Qt.ItemFlag.ItemIsAutoTristate:
                            model.dataChanged.emit(parent.index(),parent.index(),[Qt.ItemDataRole.CheckStateRole])
                        else: break
        super().setData(value, role)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec())