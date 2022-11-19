import XLSX from 'xlsx';
import fs from 'fs';
import * as excel_config from './excel_config.js';
import * as utils from './utils.js';

// Excel data
const config = excel_config.config;
const workbook = XLSX.readFile('./graph.xlsx');
const schedule_sheet = workbook.Sheets[config.schedule_sheet_name];
const sites_sheet = workbook.Sheets[config.sites_sheet_name];

function schedule_map() {
   let map = []
   for (let r = config.names_id[0].r; r <= config.names_id[1].r; r++) {

      let row = {
         i: map.length,
         name: schedule_sheet[XLSX.utils.encode_cell({ r: r, c: config.days_id[0].c - 1 })].v,
         data: [],
         month_hours: 0
      }

      for (let c = config.days_id[0].c; c <= config.days_id[1].c; c++) {
         let cell_id = XLSX.utils.encode_cell({ r: r, c: c })
         let day = {
            data: schedule_sheet[cell_id]?.v ?? '',
            cell: cell_id
         }
         day.data = day.data?.toString()?.toUpperCase()
         day.data = day.data == config.dayoff || 
                                day.data == config.vacation || 
                                day.data == config.misc_dayoff 
                     ? 'NA' : '' // change
         row.data.push(day)
      }
      map.push(row)
   }

   return map
}

function sites_data() {
   let sites = []

   for (let r = config.hours_id[0].r; r <= config.hours_id[1].r; r++) {
      let hours = sites_sheet[XLSX.utils.encode_cell({ r: r, c: config.hours_id[0].c + 1 })]?.v
         ?.split('-')
         ?.map(num => num.replace(" (суббота)", ""))
         ?.filter(num => utils.isNumeric(num) && Number(num) > 0)
         ?.map(num => Number(num));
      if (!hours) continue;

      let days = [
         sites_sheet[XLSX.utils.encode_cell({ r: r, c: config.hours_id[0].c + 3 })].v,
         sites_sheet[XLSX.utils.encode_cell({ r: r, c: config.hours_id[0].c + 4 })].v
      ]

      let work_day = hours[1] - hours[0];

      let name = sites_sheet[XLSX.utils.encode_cell({ r: r, c: config.hours_id[0].c })].v
      let id = `${name}_${sites.length}`
      let max_workers = sites_sheet[XLSX.utils.encode_cell({ r: r, c: config.hours_id[0].c + 5 })].v

      let site = {
         id: id,
         name: name,
         work_day: work_day,
         hours: hours,
         days: days,
         shift: name.split('/')[1] ?? '',
         max_workers: max_workers,
      }
      sites.push(site)
   }
   return sites
}

function select_site(worker, day_num, sites_days, ignore_restrictions) {
   if (ignore_restrictions) {
      for (let s of worker.total_on_sites) {
         if (sites_days[day_num][s.site_data.id].max > 0)
            return s;
      }
   }
   let site = worker.total_on_sites.reduce(function (prev, curr) {
      const shift = (worker.data[day_num].data !== curr.site_data.shift) && (worker.data[day_num].data !== '') // Проверка на смену
      return (
         (prev.total < curr.total) || // Та точка где работник провёл меньше всего времени
         (shift) || // Смена
         ((sites_days[day_num][prev.site_data.id].current == 0) && (prev.total <= curr.total)) // Макс число рыл на точку
      ) && sites_days[day_num][prev.site_data.id].max != 0
         ? prev : curr;
   });
   if (sites_days[day_num][site.site_data.id].max == 0)
      return null;
   return site;
}

function create_schedule() {
   let schedule = schedule_map()
   let sites = sites_data();
   let sites_days = schedule[0].data.map((_, i) => sites.reduce(function (map, site) {
      let day = (i + config.starting_day - 1) % 7 + 1
      map[site.id] = {
         max: site.days[0] <= day && site.days[1] >= day ? site.max_workers : 0,
         current: 0,
         active: site.days[0] <= day && site.days[1] >= day,
         day: day,
         days: site.days
      };
      return map;
   }, {}));

   schedule.map(worker => {
      worker.total_on_sites = sites.map(site => { return { site_data: site, total: 0 } })
   })

   // главная часть
   let total_days = config.days_id[1].c - config.days_id[0].c + 1
   for (let i = 0; i < total_days; i++) {
      for (let k = 1; k <= 3; k++) { // первый проход по предпочтениям, второй основной, третий остаточный
         let workers_sorted = schedule.sort((a, b) => a.month_hours - b.month_hours).map(w => w.i)

         for (let j = 0; j < workers_sorted.length; j++) {
            if (k == 1 && schedule[j].data[i].data === '') continue; // Сначала предпочтения по сменам 
            if (schedule[j].data[i].data == 'NA') {
               schedule[j].data[i] = 'ПРОПУСК'
               continue
            }
            if (schedule[j].data[i].data === undefined)
               continue;
            let site = select_site(schedule[j], i, sites_days, k == 3)
            if (site === null)
               break;
            schedule[j].data[i] = site.site_data.id;
            site.total++;
            sites_days[i][site.site_data.id].max--;
            sites_days[i][site.site_data.id].current++;
            schedule[j].month_hours += site.site_data.work_day
         }
      }
   }

   // Очистка
   schedule.map(w => {
      delete w.total_on_sites;
      for (let i = 0; i < w.data.length; i++)
         if (w.data[i].cell)
            w.data[i] = null
      return w;
   })
   schedule.sort((a, b) => a.i - b.i).map(w => w.i)

   return schedule;
}

function add_schedule_to_xlsx(schedule) {
   const res_file_name = './result.xlsx';
   fs.copyFile(config.file_name, res_file_name);

   const new_workbook = XLSX.readFile(res_file_name);

   for (let r = config.names_id[0].r; r <= config.names_id[1].r; r++) {
      let i = r - config.names_id[0].r;
      for (let c = config.days_id[0].c; c <= config.days_id[1].c; c++) {
         let j = c - config.days_id[0].c;
         let cell_id = XLSX.utils.encode_cell({ r: r, c: c })

         new_workbook.Sheets[config.schedule_sheet_name][cell_id] = {
               h: schedule[i].data[j],
               r: `<t>${schedule[i].data[j]}</t>`,
               t: 's',
               v: schedule[i].data[j],
               w: schedule[i].data[j],
         };
      }
   }
   XLSX.writeFile(new_workbook, res_file_name) // write the same workbook with new values
}

const schedule = create_schedule();
add_schedule_to_xlsx(schedule)
