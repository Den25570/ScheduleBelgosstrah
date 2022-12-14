import * as XLSX from 'xlsx';

export const config = {
    // Путь к excel файлу
    file_name: './graph.xlsx', 
    // Основной лист с расписанием
    schedule_sheet_name: "01 График",
    // Лист с точками
    sites_sheet_name: "Приложение",
    // Лист с приложением к СУ
    sites_hours_sheet_name: "Приложение к СУ",
    // День недели с которого начинается расписание (1-7) (Пон. - вс.)
    schedule_starting_day: 4,
    // Строки в которых находятся имена работников из основного листа
    names_id: [
        XLSX.utils.decode_cell('B11'),
        XLSX.utils.decode_cell('B56'),
    ],
    // Столбцы в которых находятся дни расписания из основного листа
    days_id: [
        XLSX.utils.decode_cell('C10'),
        XLSX.utils.decode_cell('AG10')
    ],
    // Строки в которых находятся условные обозначения точек из листа с точками (Приложение)
    hours_id: [
        XLSX.utils.decode_cell('A12'),
        XLSX.utils.decode_cell('A40'),
    ],
    // Cимвол выходного 
    dayoff: 'В',
    // Символ отпуска
    vacation: 'О',
    // Символ для прочего
    misc_dayoff: 'СУ'
};

