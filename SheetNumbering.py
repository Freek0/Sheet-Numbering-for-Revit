# -*- coding: utf-8 -*-
"""
Нумерация листов в Revit с графическим интерфейсом
Показывает окно со списком всех листов, позволяет выбрать нужные и указать начальный номер
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('PresentationFramework')

from Autodesk.Revit import DB
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from RevitServices import Elements
from System.Windows import Application, Window
from System.Windows.Controls import CheckBox, Button, TextBox, Label, ScrollViewer, StackPanel, DockPanel, Grid, GridSplitter, ComboBox
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Media import Brushes
from System.Windows.Input import Keyboard, ModifierKeys
import System

# Получаем текущий документ
doc = DocumentManager.Instance.CurrentDBDocument

# Получаем все листы
sheet_collector = DB.FilteredElementCollector(doc)
sheet_collector.OfClass(DB.ViewSheet)
all_sheets = sheet_collector.ToElements()

# Фильтруем только активные листы
sheets_list = []
for sheet in all_sheets:
    if sheet is not None and not sheet.IsPlaceholder:
        sheets_list.append(sheet)

# Функция для естественной сортировки (natural sort) - как в Project Browser Revit
# Преобразует строку в кортеж с типами для правильного сравнения смешанных типов
import re
def natural_sort_key(text):
    """
    Преобразует строку в ключ для естественной сортировки.
    Возвращает кортеж кортежей: ((тип, значение), ...)
    где тип: 0 для чисел, 1 для строк
    Это позволяет корректно сравнивать смешанные значения.
    Пример: '12' -> ((0, 12),), '2' -> ((0, 2),), 'A10' -> ((1, 'a'), (0, 10))
    Сортирует так же, как Project Browser в Revit (1, 2, 3, ..., 9, 10, 11, ...)
    """
    if text is None:
        return ((1, ''),)
    text = str(text)
    # Разбиваем строку на части: числа и нечисловые символы
    parts = []
    for part in re.split(r'(\d+)', text):
        if part:
            # Если часть - число, преобразуем в int, иначе оставляем строкой
            try:
                parts.append((0, int(part)))  # 0 - тип для чисел (сравниваются как числа)
            except ValueError:
                parts.append((1, part.lower()))  # 1 - тип для строк (сравниваются без учета регистра)
    return tuple(parts) if parts else ((1, text.lower()),)

# Функция для получения значения параметра из листа
def get_sheet_parameter_value(sheet, param_name):
    """Получает значение параметра из листа по имени"""
    try:
        param = sheet.LookupParameter(param_name)
        if param is not None:
            storage_type = param.StorageType
            if storage_type == DB.StorageType.String:
                value = param.AsString()
            elif storage_type == DB.StorageType.Integer:
                value = param.AsInteger()
            elif storage_type == DB.StorageType.Double:
                value = param.AsDouble()
            elif storage_type == DB.StorageType.ElementId:
                elem_id = param.AsElementId()
                if elem_id.IntegerValue >= 0:
                    elem = doc.GetElement(elem_id)
                    value = elem.Name if elem else None
                else:
                    value = None
            else:
                value = param.AsValueString()
            return value if value is not None else ""
        return ""
    except:
        return ""

# Сохраняем полный список листов для фильтрации
all_sheets_list = list(sheets_list)

# Получаем уникальные значения параметра "ADSK_Штамп Раздел проекта"
parameter_name = "ADSK_Штамп Раздел проекта"
parameter_values = set()
has_empty_values = False
for sheet in all_sheets_list:
    param_value = get_sheet_parameter_value(sheet, parameter_name)
    param_str = str(param_value) if param_value else ""
    if param_str.strip():
        parameter_values.add(param_str)
    else:
        has_empty_values = True
parameter_values_list = sorted(list(parameter_values))
if has_empty_values:
    parameter_values_list.append("(Без значения)")

# Сортируем по номеру листа с использованием естественной сортировки
sheets_list.sort(key=lambda x: natural_sort_key(x.SheetNumber))

if len(sheets_list) == 0:
    OUT = "Ошибка: В документе нет листов для нумерации"
else:
    # Используем список для хранения результата (вместо nonlocal)
    result_data = {'dialog_result': False, 'selected_sheets': [], 'start_number': 1, 'prefix': ''}
    checkboxes = []
    all_checkboxes = []  # Храним все чекбоксы для фильтрации
    ui_state = {'last_checked_index': -1}  # Индекс последнего выбранного чекбокса для Shift-выбора
    
    # Основное окно
    window = Window()
    window.Title = "Нумерация листов"
    window.Width = 900
    window.Height = 700
    window.MinWidth = 600
    window.MinHeight = 500
    window.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
    window.ResizeMode = System.Windows.ResizeMode.CanResize
    
    # Основной контейнер - используем Grid для лучшего контроля
    main_grid = Grid()
    main_grid.Margin = Thickness(10)
    
    # Создаем строки: верх (фильтр), верх (кнопки), средняя часть (растягиваемая), низ
    main_grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition())
    main_grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition())
    main_grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition())
    main_grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition())
    main_grid.RowDefinitions[0].Height = System.Windows.GridLength(0, System.Windows.GridUnitType.Auto)  # Фильтр
    main_grid.RowDefinitions[1].Height = System.Windows.GridLength(0, System.Windows.GridUnitType.Auto)  # Кнопки
    main_grid.RowDefinitions[2].Height = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)  # Список
    main_grid.RowDefinitions[3].Height = System.Windows.GridLength(0, System.Windows.GridUnitType.Auto)  # Низ
    
    # Панель фильтра - дропдаун для выбора раздела проекта
    filter_panel = StackPanel()
    filter_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
    filter_panel.Margin = Thickness(0, 0, 0, 10)
    
    filter_label = Label()
    filter_label.Content = "Раздел проекта:"
    filter_label.Margin = Thickness(0, 0, 10, 0)
    filter_label.VerticalAlignment = VerticalAlignment.Center
    
    filter_combo = ComboBox()
    filter_combo.Width = 250
    filter_combo.Height = 25
    filter_combo.VerticalAlignment = VerticalAlignment.Center
    filter_combo.Items.Add("Все")
    for value in parameter_values_list:
        filter_combo.Items.Add(value)
    filter_combo.SelectedIndex = 0  # По умолчанию "Все"
    
    filter_panel.Children.Add(filter_label)
    filter_panel.Children.Add(filter_combo)
    Grid.SetRow(filter_panel, 0)
    main_grid.Children.Add(filter_panel)
    
    # Верхняя панель - управление выбором
    top_panel = StackPanel()
    top_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
    top_panel.Margin = Thickness(0, 0, 0, 10)
    
    btn_select_all = Button()
    btn_select_all.Content = "Выбрать все"
    btn_select_all.Margin = Thickness(0, 0, 5, 0)
    btn_select_all.Width = 100
    
    btn_deselect_all = Button()
    btn_deselect_all.Content = "Снять все"
    btn_deselect_all.Margin = Thickness(0, 0, 5, 0)
    btn_deselect_all.Width = 100
    
    top_panel.Children.Add(btn_select_all)
    top_panel.Children.Add(btn_deselect_all)
    Grid.SetRow(top_panel, 1)
    main_grid.Children.Add(top_panel)
    
    # Средняя часть - список листов с прокруткой (заполняет всю ширину, изменяется при изменении размера окна)
    scroll_viewer = ScrollViewer()
    scroll_viewer.Margin = Thickness(0, 0, 0, 10)
    scroll_viewer.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
    scroll_viewer.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
    
    sheets_panel = StackPanel()
    sheets_panel.Orientation = System.Windows.Controls.Orientation.Vertical
    
    # Создаем чекбоксы для каждого листа
    def checkbox_preview_mousedown(sender, e):
        try:
            # Если зажат Shift, обрабатываем выбор диапазона
            if e.ChangedButton == System.Windows.Input.MouseButton.Left:
                if Keyboard.Modifiers == ModifierKeys.Shift and ui_state['last_checked_index'] >= 0:
                    current_index = checkboxes.index(sender)
                    # Выбираем диапазон от last_checked_index до current_index
                    start_idx = min(ui_state['last_checked_index'], current_index)
                    end_idx = max(ui_state['last_checked_index'], current_index)
                    
                    # Запоминаем состояние последнего чекбокса
                    last_state = checkboxes[ui_state['last_checked_index']].IsChecked
                    
                    # Выбираем/снимаем все в диапазоне
                    for idx in range(start_idx, end_idx + 1):
                        checkboxes[idx].IsChecked = last_state
                    
                    ui_state['last_checked_index'] = current_index
                    e.Handled = True
        except:
            pass
    
    def checkbox_click(sender, e):
        try:
            # Запоминаем индекс последнего выбранного чекбокса
            current_index = checkboxes.index(sender)
            ui_state['last_checked_index'] = current_index
        except:
            pass
    
    # Функция для создания чекбокса
    def create_checkbox(sheet):
        checkbox = CheckBox()
        checkbox.Content = "[{0}] {1}".format(sheet.SheetNumber, sheet.Name)
        checkbox.Tag = sheet
        checkbox.IsChecked = False
        checkbox.Margin = Thickness(5)
        checkbox.PreviewMouseDown += checkbox_preview_mousedown
        checkbox.Click += checkbox_click
        return checkbox
    
    # Создаем чекбоксы для всех листов (сохраняем в all_checkboxes)
    for sheet in all_sheets_list:
        checkbox = create_checkbox(sheet)
        all_checkboxes.append(checkbox)
        checkboxes.append(checkbox)
        sheets_panel.Children.Add(checkbox)
    
    # Функция для фильтрации списка листов
    def filter_sheets(sender, e):
        try:
            selected_value = filter_combo.SelectedItem
            if selected_value is None:
                return
            
            # Очищаем текущий список
            sheets_panel.Children.Clear()
            checkboxes.clear()
            
            # Если выбрано "Все", показываем все листы
            if selected_value == "Все":
                filtered_sheets = all_sheets_list
            elif selected_value == "(Без значения)":
                # Показываем листы без значения параметра
                filtered_sheets = []
                for sheet in all_sheets_list:
                    param_value = get_sheet_parameter_value(sheet, parameter_name)
                    param_str = str(param_value) if param_value else ""
                    if not param_str.strip():
                        filtered_sheets.append(sheet)
            else:
                # Фильтруем по значению параметра
                filtered_sheets = []
                for sheet in all_sheets_list:
                    param_value = get_sheet_parameter_value(sheet, parameter_name)
                    param_str = str(param_value) if param_value else ""
                    if param_str == selected_value:
                        filtered_sheets.append(sheet)
                
                # Сортируем отфильтрованный список
                filtered_sheets.sort(key=lambda x: natural_sort_key(x.SheetNumber))
            
            # Создаем чекбоксы для отфильтрованных листов
            for sheet in filtered_sheets:
                # Находим соответствующий чекбокс из всех
                for cb in all_checkboxes:
                    if cb.Tag == sheet:
                        checkbox = cb
                        break
                else:
                    # Если не нашли, создаем новый
                    checkbox = create_checkbox(sheet)
                    all_checkboxes.append(checkbox)
                
                checkboxes.append(checkbox)
                sheets_panel.Children.Add(checkbox)
                
            # Сбрасываем индекс последнего выбранного
            ui_state['last_checked_index'] = -1
        except Exception as ex:
            pass
    
    # Подключаем обработчик изменения фильтра
    filter_combo.SelectionChanged += filter_sheets
    
    scroll_viewer.Content = sheets_panel
    Grid.SetRow(scroll_viewer, 2)
    main_grid.Children.Add(scroll_viewer)
    
    # Нижняя панель - начальный номер и кнопки (закреплены справа внизу)
    bottom_grid = Grid()
    bottom_grid.Margin = Thickness(0, 10, 0, 0)
    
    # Создаем колонки для Grid: левая часть растягивается, правая - авторазмер
    bottom_grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())
    bottom_grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())
    bottom_grid.ColumnDefinitions[0].Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
    bottom_grid.ColumnDefinitions[1].Width = System.Windows.GridLength(0, System.Windows.GridUnitType.Auto)
    
    # Левая часть - префикс и начальный номер
    left_panel = StackPanel()
    left_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
    left_panel.VerticalAlignment = VerticalAlignment.Center
    
    prefix_label = Label()
    prefix_label.Content = "Префикс:"
    prefix_label.Margin = Thickness(0, 0, 10, 0)
    prefix_label.VerticalAlignment = VerticalAlignment.Center
    
    prefix_box = TextBox()
    prefix_box.Text = ""
    prefix_box.Width = 80
    prefix_box.VerticalAlignment = VerticalAlignment.Center
    prefix_box.ToolTip = "Префикс перед номером (например, A, 1-A, и т.д.)"
    
    start_label = Label()
    start_label.Content = "Начальный номер:"
    start_label.Margin = Thickness(20, 0, 10, 0)
    start_label.VerticalAlignment = VerticalAlignment.Center
    
    start_number_box = TextBox()
    start_number_box.Text = "1"
    start_number_box.Width = 60
    start_number_box.VerticalAlignment = VerticalAlignment.Center
    
    left_panel.Children.Add(prefix_label)
    left_panel.Children.Add(prefix_box)
    left_panel.Children.Add(start_label)
    left_panel.Children.Add(start_number_box)
    Grid.SetColumn(left_panel, 0)
    bottom_grid.Children.Add(left_panel)
    
    # Правая часть - кнопки (закреплены справа)
    right_panel = StackPanel()
    right_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
    right_panel.HorizontalAlignment = HorizontalAlignment.Right
    right_panel.VerticalAlignment = VerticalAlignment.Center
    
    btn_ok = Button()
    btn_ok.Content = "Выполнить нумерацию"
    btn_ok.Width = 150
    btn_ok.Height = 30
    btn_ok.Margin = Thickness(0, 0, 10, 0)
    
    btn_cancel = Button()
    btn_cancel.Content = "Отмена"
    btn_cancel.Width = 80
    btn_cancel.Height = 30
    
    right_panel.Children.Add(btn_ok)
    right_panel.Children.Add(btn_cancel)
    Grid.SetColumn(right_panel, 1)
    bottom_grid.Children.Add(right_panel)
    
    Grid.SetRow(bottom_grid, 3)
    main_grid.Children.Add(bottom_grid)
    
    # Обработчики событий
    def select_all(sender, e):
        for cb in checkboxes:
            cb.IsChecked = True
    
    def deselect_all(sender, e):
        for cb in checkboxes:
            cb.IsChecked = False
    
    def ok_click(sender, e):
        result_data['selected_sheets'] = []
        # Собираем выбранные листы из всех чекбоксов, а не только видимых
        for cb in all_checkboxes:
            if cb.IsChecked == True:
                result_data['selected_sheets'].append(cb.Tag)
        try:
            result_data['start_number'] = int(start_number_box.Text)
        except:
            result_data['start_number'] = 1
        result_data['prefix'] = prefix_box.Text.strip() if prefix_box.Text else ''
        result_data['dialog_result'] = True
        window.DialogResult = True
        window.Close()
    
    def cancel_click(sender, e):
        window.DialogResult = False
        window.Close()
    
    btn_select_all.Click += select_all
    btn_deselect_all.Click += deselect_all
    btn_ok.Click += ok_click
    btn_cancel.Click += cancel_click
    
    window.Content = main_grid
    
    # Запускаем окно
    result = window.ShowDialog()
    
    # Обрабатываем результат
    if result == True and result_data['dialog_result'] == True and len(result_data['selected_sheets']) > 0:
        try:
            # Начинаем транзакцию
            TransactionManager.Instance.ForceCloseTransaction()
            t = DB.Transaction(doc, "Нумерация листов")
            t.Start()
            
            renumbered_sheets = []
            current_number = result_data['start_number']
            errors = []
            
            # Нумеруем выбранные листы
            for sheet in result_data['selected_sheets']:
                if hasattr(sheet, 'InternalElement'):
                    sheet_unwrapped = sheet.InternalElement
                else:
                    sheet_unwrapped = sheet
                
                if sheet_unwrapped is not None and hasattr(sheet_unwrapped, 'SheetNumber'):
                    try:
                        old_number = sheet_unwrapped.SheetNumber
                        # Формируем новый номер: префикс + номер
                        prefix = result_data.get('prefix', '')
                        if prefix:
                            new_number = prefix + str(current_number)
                        else:
                            new_number = str(current_number)
                        
                        sheet_unwrapped.SheetNumber = new_number
                        renumbered_sheets.append("Лист {0} -> {1}".format(old_number, new_number))
                        current_number += 1
                    except Exception as e:
                        errors.append("Ошибка при нумерации листа: {0}".format(str(e)))
            
            t.Commit()
            
            # Формируем результат
            result_parts = ["Нумерация завершена!"]
            prefix_str = result_data.get('prefix', '')
            if prefix_str:
                result_parts.append("Префикс: {0}".format(prefix_str))
            result_parts.append("Начальный номер: {0}".format(result_data['start_number']))
            result_parts.append("Обработано листов: {0}".format(len(renumbered_sheets)))
            
            if len(renumbered_sheets) > 0:
                result_parts.append("")
                for i, result in enumerate(renumbered_sheets[:50]):
                    result_parts.append(result)
                if len(renumbered_sheets) > 50:
                    result_parts.append("... и еще {0} листов".format(len(renumbered_sheets) - 50))
            
            if len(errors) > 0:
                result_parts.append("")
                result_parts.append("Ошибки:")
                for error in errors:
                    result_parts.append(error)
            
            OUT = "\r\n".join(result_parts)
        except Exception as e:
            import traceback
            error_msg = "Ошибка при выполнении скрипта:\r\n{0}\r\n\r\n{1}".format(str(e), traceback.format_exc())
            OUT = error_msg
    elif result == False:
        OUT = "Операция отменена пользователем"
    else:
        OUT = "Ошибка: Не выбраны листы для нумерации"

