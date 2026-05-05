import React, { useState, useCallback } from 'react';
import * as XLSXRead from 'xlsx';
import XLSX from 'xlsx-js-style';
import { 
  Trash2,
  FileText, 
  Upload, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  FileSpreadsheet,
  Bus,
  Map,
  History,
  Calendar,
  List,
  Lock,
  KeyRound,
  User
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- Utilities ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

type MainFileType = 'registry_mun' | 'registry_inter';
type AppTab = 'main' | 'decoder';

interface FileState {
  file: File | null;
  data: any[] | null;
  status: 'empty' | 'loading' | 'success' | 'processed' | 'error';
}

const FILE_CONFIG: Record<MainFileType, { label: string; icon: any; description: string }> = {
  registry_mun: { 
    label: 'Реестр муниципальный', 
    icon: Bus, 
    description: 'Данные по городским маршрутам' 
  },
  registry_inter: { 
    label: 'Реестр межмуниципальный', 
    icon: Map, 
    description: 'Данные по пригородным маршрутам' 
  },
};

const SCHEDULE_TYPES: Record<string, string> = {
  '21': 'Зима (будни)',
  '24': 'Зима (выходные)',
  '31': 'Лето (будни)',
  '34': 'Лето (выходные)',
};

// --- Components ---
interface FileCardProps {
  type: string;
  config: { label: string; icon: any; description: string };
  state: FileState;
  onUpload: (file: File) => void;
  onRemove: () => void;
  accept?: string;
}

const FileCard: React.FC<FileCardProps> = ({ 
  type, 
  config, 
  state, 
  onUpload,
  onRemove,
  accept
}) => {
  const [isOver, setIsOver] = useState(false);

  return (
    <div className="bg-white p-4 rounded-2xl shadow-sm transition-all flex flex-col gap-4 group relative">
      <div className="flex justify-between items-center px-2">
        <h4 className="text-[11px] font-bold uppercase tracking-wide text-[#1A237E]">
          {config.label}
        </h4>
        {state.status === 'success' && (
          <button onClick={onRemove} className="text-slate-300 hover:text-[#D43A3A] transition-colors"><Trash2 className="w-4 h-4" /></button>
        )}
      </div>

      <div 
        onDragOver={(e) => { e.preventDefault(); setIsOver(true); }}
        onDragLeave={() => setIsOver(false)}
        onDrop={(e) => { e.preventDefault(); setIsOver(false); const file = e.dataTransfer.files[0]; if (file) onUpload(file); }}
        className={cn(
          "relative h-24 rounded-xl flex flex-col justify-center items-center gap-[6px] px-4 transition-all overflow-hidden",
          state.status === 'success' ? "bg-[#E5F5EF] border border-transparent" : "bg-[#F0F4F8] border border-transparent",
          isOver && "ring-2 ring-[#00A478] ring-offset-2"
        )}
      >
        <div className={cn("rounded-lg flex items-center justify-center", state.status === 'success' ? "text-[#00A478]" : "text-[#94A3B8]")}>
          <config.icon className="w-6 h-6" />
        </div>
        
        <div className="flex flex-col items-center min-w-0 text-center">
          <span className="text-[11px] font-bold text-[#1A237E] truncate w-full uppercase tracking-tight px-2">
            {state.file ? state.file.name : 'ФАЙЛ НЕ ВЫБРАН'}
          </span>
          <span className={cn("text-[9px] font-bold uppercase tracking-widest mt-1", (state.status === 'success' || state.status === 'processed') ? "text-[#00A478]" : "text-[#94A3B8]")}>
            {state.status === 'loading' && 'ЧТЕНИЕ ДАННЫХ...'}
            {state.status === 'success' && 'ФАЙЛ ГОТОВ'}
            {state.status === 'processed' && 'ДАННЫЕ ИМПОРТИРОВАНЫ'}
            {state.status === 'error' && 'ОШИБКА ЧТЕНИЯ'}
            {state.status === 'empty' && 'ОЖИДАНИЕ ВВОДА'}
          </span>
        </div>

        <input 
          type="file" 
          accept={accept}
          className="absolute inset-0 opacity-0 cursor-pointer"
          onChange={(e) => { const file = e.target.files?.[0]; if (file) onUpload(file); }}
        />
        {state.status !== 'success' && <Upload className="absolute top-3 right-3 w-4 h-4 text-[#94A3B8] opacity-50" />}
      </div>

      <button className={cn(
        "text-[12px] font-bold uppercase tracking-widest py-3 rounded-xl border transition-all text-center",
        state.status === 'success' ? "bg-[#00A478] text-white border-transparent cursor-default shadow-md" : "bg-white text-[#1A237E] border-[#DEE4ED] hover:bg-slate-50 active:scale-95"
      )}>
        {state.status === 'success' ? 'ФАЙЛ ПРИНЯТ' : 'ВЫБРАТЬ ФАЙЛ'}
      </button>
    </div>
  );
}

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(() => localStorage.getItem('auth') === 'true');
  const [password, setPassword] = useState('');
  const [authError, setAuthError] = useState(false);
  
  const [activeTab, setActiveTab] = useState<AppTab>('decoder');
  const [files, setFiles] = useState<Record<MainFileType, FileState>>({
    registry_mun: { file: null, data: null, status: 'empty' },
    registry_inter: { file: null, data: null, status: 'empty' },
  });

  const [winterDecoderFile, setWinterDecoderFile] = useState<FileState>({ file: null, data: null, status: 'empty' });
  const [summerDecoderFile, setSummerDecoderFile] = useState<FileState>({ file: null, data: null, status: 'empty' });
  const [accumulatorReportFile, setAccumulatorReportFile] = useState<FileState>({ file: null, data: null, status: 'empty' });
  const [targetReportFile, setTargetReportFile] = useState<FileState>({ file: null, data: null, status: 'empty' });
  const [decoderWorkbook, setDecoderWorkbook] = useState<any | null>(null);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [resultWorkbook, setResultWorkbook] = useState<any | null>(null);
  const [resultSummary, setResultSummary] = useState<{ route: string; stopsCount: number; registryMatch: boolean } | null>(null);
  const [currentRoute, setCurrentRoute] = useState<string | null>(null);

  const extractDigits = (s: any) => {
    if (s === undefined || s === null) return "";
    let str = String(s).trim();
    // Эквивалент Pandas str.extract('(\d+)')
    // Удаляем .0 (артефакт Excel)
    if (str.endsWith('.0')) str = str.slice(0, -2);
    
    const match = str.match(/\d+/);
    return match ? match[0] : "";
  };

  const cleanRouteData = (s: any) => {
    if (s === undefined || s === null) return "";
    // Для первых трех колонок: удаляем кавычки, скобки и точки с запятой
    return String(s).replace(/[";()]/g, '').trim();
  };

  const cleanGeneral = (s: any) => {
    if (s === undefined || s === null) return "";
    // Для всего остального файла: удаляем только кавычки
    return String(s).replace(/["]/g, '').trim();
  };

  const normalizeString = (s: any) => {
    if (s === undefined || s === null) return "";
    // Для сравнения в реестре: удаляем кавычки и пробелы, приводим к нижнему регистру
    return String(s).replace(/["]/g, '').trim().toLowerCase();
  };

  const findRegistryEntry = (routeId: string, allRows: any[]) => {
    // В итоговом файле в колонке №3 примени жесткую очистку: удали все кроме цифр
    const targetDigits = extractDigits(routeId);
    if (!targetDigits) return null;

    // Сравниваем с номером маршрута из Реестра (индекс 1)
    return allRows.find(row => {
      const regRouteVal = row['__col_1'];
      // Для реестра тоже применяем extractDigits, чтобы найти соответствие с "грязным" номером
      return extractDigits(regRouteVal) === targetDigits;
    });
  };

  const handleFileChange = useCallback((type: MainFileType, file: File) => {
    setFiles(prev => ({ ...prev, [type]: { ...prev[type], status: 'loading', file } }));
    setResultWorkbook(null);
    setResultSummary(null);
    setError(null);

    // Use setTimeout to allow UI to update with 'loading' status before heavy sync work
    setTimeout(async () => {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const isCsv = file.name.toLowerCase().endsWith('.csv');
        const workbook = XLSXRead.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        let rawData: any[][] = XLSXRead.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        
        if (isCsv && rawData.length > 0 && rawData[0].length === 1 && String(rawData[0][0]).includes(';')) {
           rawData = rawData.map(row => String(row[0]).split(';'));
        }

        const keysToSearch = ["порядковыйномер", "наименованиемаршрута", "номермаршрута", "регистрационныйномер"];
        
        let bestHeaderIndex = 0;
        let maxMatches = -1;
        for (let i = 0; i < Math.min(rawData.length, 30); i++) {
          const row = rawData[i];
          if (!row || !Array.isArray(row)) continue;
          let rowMatches = 0;
          row.forEach(cell => {
            const norm = normalizeString(cell);
            if (keysToSearch.some(k => norm.includes(k))) rowMatches++;
          });
          if (rowMatches > maxMatches) {
            maxMatches = rowMatches;
            bestHeaderIndex = i;
          }
        }

        const headerRowIndex = bestHeaderIndex;
        const jsonData = rawData.slice(headerRowIndex + 1).filter(row => {
          const cell0 = String(row[0] || "").trim();
          const cell1 = String(row[1] || "").trim();
          const cell2 = String(row[2] || "").trim();
          if (cell0 === "1" && cell1 === "2" && cell2 === "3") return false;
          const firstColsText = row.slice(0, 5).join(' ').toLowerCase();
          if (firstColsText.includes("класс") || firstColsText.includes("вместимость")) return false;
          return row.some(cell => cell !== "" && cell !== null && cell !== undefined);
        }).map((row) => {
          return {
            '__col_0': cleanGeneral(row[0]),
            '__col_1': cleanGeneral(row[1]),
            '__col_2': cleanGeneral(row[2])
          };
        });

        setFiles(prev => ({
          ...prev,
          [type]: { file, data: jsonData, status: 'success' }
        }));
      } catch (err) {
        console.error(err);
        setFiles(prev => ({ ...prev, [type]: { ...prev[type], status: 'error' } }));
        setError("Ошибка при обработке файла реестра");
      }
    }, 10);
  }, []);

  const handleAccumulatorReportFile = useCallback((file: File) => {
    setAccumulatorReportFile({ file, data: null, status: 'loading' });
    setError(null);

    setTimeout(async () => {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSXRead.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        let rawData: any[][] = XLSXRead.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        
        const isCsv = file.name.toLowerCase().endsWith('.csv');
        if (isCsv && rawData.length > 0 && rawData[0].length === 1 && String(rawData[0][0]).includes(';')) {
           rawData = rawData.map(row => String(row[0]).split(';'));
        }

        setAccumulatorReportFile({ file, data: rawData, status: 'success' });
      } catch (err) {
        console.error(err);
        setAccumulatorReportFile({ file: null, data: null, status: 'error' });
        setError("Ошибка при чтении Накопительного отчета");
      }
    }, 10);
  }, []);

  const handleWinterDecoderFile = useCallback((file: File) => {
    setWinterDecoderFile({ file, data: null, status: 'loading' });
    setDecoderWorkbook(null);
    setCurrentRoute(null);
    setError(null);
    
    setTimeout(async () => {
      try {
        const text = await file.text();
        const lines = text.split(/\r?\n/).filter(line => line.trim().includes('_'));
        if (lines.length === 0) {
          throw new Error('В файле не найдено подходящих данных (строки должны содержать "_")');
        }
        const firstRoute = lines[0].split('_')[0].trim();
        setCurrentRoute(firstRoute);
        setWinterDecoderFile({ file, data: lines, status: 'success' });
      } catch (err: any) {
        setError(err.message || "Ошибка при чтении файла");
        setWinterDecoderFile(prev => ({ ...prev, status: 'error' }));
      }
    }, 10);
  }, []);

  const handleSummerDecoderFile = useCallback((file: File) => {
    setSummerDecoderFile({ file, data: null, status: 'loading' });
    setDecoderWorkbook(null);
    setError(null);
    
    setTimeout(async () => {
      try {
        const text = await file.text();
        const lines = text.split(/\r?\n/).filter(line => line.trim().includes('_'));
        if (lines.length === 0) {
          throw new Error('В файле не найдено подходящих данных (строки должны содержать "_")');
        }
        setSummerDecoderFile({ file, data: lines, status: 'success' });
      } catch (err: any) {
        setError(err.message || "Ошибка при чтении файла");
        setSummerDecoderFile(prev => ({ ...prev, status: 'error' }));
      }
    }, 10);
  }, []);

  const runDecoder = () => {
    if ((!winterDecoderFile.data || winterDecoderFile.data.length === 0) && (!summerDecoderFile.data || summerDecoderFile.data.length === 0)) return;
    
    try {
      const stopGroups: Record<string, { name: string; code: string; times: Set<string> }> = {};
      
      const processLineDecoder = (line: string, season: string) => {
        const parts = line.split('_');
        if (parts.length < 9) return;

        const route = parts[0];
        const type = SCHEDULE_TYPES[parts[1]] || parts[1] || season;
        const timeDep = parts[4];
        const stopCode = parts[6];
        const stopName = parts[8];

        const groupKey = `${route}_${type}_${stopCode}_${stopName}`;

        if (!stopGroups[groupKey]) {
          stopGroups[groupKey] = {
            name: stopName,
            code: stopCode,
            times: new Set()
          };
        }
        if (timeDep) stopGroups[groupKey].times.add(timeDep);
      };

      if (winterDecoderFile.data) winterDecoderFile.data.forEach(l => processLineDecoder(l, 'Зима'));
      if (summerDecoderFile.data) summerDecoderFile.data.forEach(l => processLineDecoder(l, 'Лето'));

      const decodedData = Object.entries(stopGroups).map(([key, data]) => {
        const [route, type] = key.split('_');
        return {
          'Маршрут': route,
          'Тип расписания': type,
          'Код остановки': data.code,
          'Наименование остановки': data.name,
          'Время отправления рейсов': Array.from(data.times).sort().join(', ')
        };
      });

      const ws = XLSX.utils.json_to_sheet(decodedData);
      ws['!cols'] = [
        { wch: 10 }, { wch: 20 }, { wch: 15 }, { wch: 30 }, { wch: 30 },
      ];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Расписание по остановкам");
      setDecoderWorkbook(wb);
    } catch (err) {
      setError("Ошибка при расшифровке данных");
    }
  };

  const [finalReportRows, setFinalReportRows] = useState<any[][] | null>(null);

  const generateReport = async () => {
    setProcessing(true);
    setError(null);
    setResultWorkbook(null);
    setFinalReportRows(null);

    try {
      if ((!winterDecoderFile.data || winterDecoderFile.data.length === 0) && (!summerDecoderFile.data || summerDecoderFile.data.length === 0)) {
        throw new Error('Необходимо загрузить хотя бы один массив данных (Зима или Лето)');
      }

      const firstLine = (winterDecoderFile.data || summerDecoderFile.data)![0];
      const routeNumberFromGtfs = firstLine.split('_')[0].trim();
      
      const allRegistryRows = [
        ...(files.registry_mun.data || []),
        ...(files.registry_inter.data || [])
      ];
      
      // Ищем соответствие в реестре по "грязному" номеру из GTFS
      const registryEntry = findRegistryEntry(routeNumberFromGtfs, allRegistryRows);
      // Очищенный номер (только цифры) для вставки в колонку №3
      const cleanRouteIdForTable = extractDigits(routeNumberFromGtfs);

      const stopsData: Record<string, { 
        winterTimes: Set<string>; 
        summerTimes: Set<string>;
        winterDays: Set<string>;
        summerDays: Set<string>;
      }> = {};
      const stopSequence: string[] = [];

      const processLine = (line: string, season: 'winter' | 'summer') => {
        const parts = line.split('_').map(p => p.trim());
        if (parts[0] === routeNumberFromGtfs) {
          const typeCode = parts[1] || "";
          const time = parts[4];
          const stopName = parts[8];
          const typeName = SCHEDULE_TYPES[typeCode] || typeCode;

          if (stopName && time) {
            if (!stopsData[stopName]) {
              stopsData[stopName] = { 
                winterTimes: new Set(), summerTimes: new Set(),
                winterDays: new Set(), summerDays: new Set()
              };
              stopSequence.push(stopName);
            }
            const currentTypeName = typeName || (season === 'winter' ? 'Зима' : 'Лето');
            const cleanTypeName = currentTypeName.includes('(') ? currentTypeName.split('(')[1].split(')')[0] : currentTypeName;
            
            if (season === 'winter') {
              stopsData[stopName].winterTimes.add(time);
              stopsData[stopName].winterDays.add(cleanTypeName);
            } else if (season === 'summer') {
              stopsData[stopName].summerTimes.add(time);
              stopsData[stopName].summerDays.add(cleanTypeName);
            }
          }
        }
      };

      if (winterDecoderFile.data) winterDecoderFile.data.forEach((line: string) => processLine(line, 'winter'));
      if (summerDecoderFile.data) summerDecoderFile.data.forEach((line: string) => processLine(line, 'summer'));

      if (!winterDecoderFile.data && summerDecoderFile.data) {
        for (const stop of stopSequence) {
          stopsData[stop].winterTimes = new Set(stopsData[stop].summerTimes);
          stopsData[stop].winterDays = new Set(Array.from(stopsData[stop].summerDays).map(d => d === 'Лето' ? 'Зима' : d));
        }
      } else if (!summerDecoderFile.data && winterDecoderFile.data) {
        for (const stop of stopSequence) {
          stopsData[stop].summerTimes = new Set(stopsData[stop].winterTimes);
          stopsData[stop].summerDays = new Set(Array.from(stopsData[stop].winterDays).map(d => d === 'Зима' ? 'Лето' : d));
        }
      }

      if (stopSequence.length === 0) {
        throw new Error(`В загруженном массиве ГТФС не найдено данных для маршрута "${routeNumberFromGtfs}".`);
      }

      const dataRows = stopSequence.map((stop: string) => {
        const row = Array(18).fill("");
        const sData = stopsData[stop];
        
        // Логика замены (Точечная) из Реестра
        // В колонку №1 (индекс 0) запиши чистый рег. номер.
        // В колонку №2 (индекс 1) запиши чистое название маршрута.
        let regNumber = registryEntry ? cleanRouteData(registryEntry['__col_0']) : "-";
        let routeName = registryEntry ? cleanRouteData(registryEntry['__col_2']) : "Наименование не найдено";
        
        // Принудительная подстраховка для №1 (если в реестре пусто)
        if (routeName === "Наименование не найдено" && cleanRouteIdForTable === "1") {
          routeName = "Железнодорожный вокзал - Ореховая улица";
        }

        // Вставляем чистые значения в первые три колонки
        row[0] = cleanRouteData(regNumber);
        row[1] = cleanRouteData(routeName);
        row[2] = cleanRouteData(cleanRouteIdForTable);
        row[3] = cleanGeneral(stop);
        row[4] = "69";
        if (sData.winterTimes.size > 0) {
          row[6] = cleanGeneral(Array.from(sData.winterDays).join(", "));
          row[7] = cleanGeneral(Array.from(sData.winterTimes).sort().join(", "));
        }
        if (sData.summerTimes.size > 0) {
          row[12] = cleanGeneral(Array.from(sData.summerDays).join(", "));
          row[13] = cleanGeneral(Array.from(sData.summerTimes).sort().join(", "));
        }
        return row;
      });

      const headerRows = [
        [
          "Регистрационный номер маршрута в реестре муниципальных, межмуниципальных, смежных межрегиональных, межрегиональных маршрутов регулярных перевозок",
          "Наименование маршрута регулярных перевозок",
          "Порядковый номер маршрута регулярных перевозок",
          "Наименования остановочных пунктов (начальные, промежуточные, конечные)",
          "Код субъекта Российской Федерации, на территории которого расположен остановочный пункт",
          "Регистрационный номер остановочного пункта в реестре остановочных пунктов по межрегиональным маршрутам регулярных перевозок (при наличии)",
          "Зимний период",
          "", "", "", "", "",
          "Летний период",
          "", "", "", "", ""
        ],
        [
          "", "", "", "", "", "",
          "Дни отправления рейсов",
          "Время отправления рейсов",
          "Время стоянки",
          "Дни прибытия рейсов",
          "Время прибытия рейсов",
          "Период действия расписания",
          "Дни отправления рейсов",
          "Время отправления рейсов",
          "Время стоянки",
          "Дни прибытия рейсов",
          "Время прибытия рейсов",
          "Период действия расписания"
        ],
        [
          "", "", "", "", "", "",
          "", "", "", "", "", "(с .... по ....)",
          "", "", "", "", "", "(с .... по ....)"
        ],
        ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18"]
      ];

      let allRows: any[][] = [];
      if (accumulatorReportFile.data && accumulatorReportFile.data.length > 0) {
        const cleanedDataRows = dataRows.map(row => row.map(cell => cleanGeneral(cell)));
        allRows = [...accumulatorReportFile.data, ...cleanedDataRows];
      } else {
        allRows = [...headerRows, ...dataRows].map(row => 
          row.map(cell => cleanGeneral(cell))
        );
      }
      const ws = XLSX.utils.aoa_to_sheet(allRows);
      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 2, c: 0 } },
        { s: { r: 0, c: 1 }, e: { r: 2, c: 1 } },
        { s: { r: 0, c: 2 }, e: { r: 2, c: 2 } },
        { s: { r: 0, c: 3 }, e: { r: 2, c: 3 } },
        { s: { r: 0, c: 4 }, e: { r: 2, c: 4 } },
        { s: { r: 0, c: 5 }, e: { r: 2, c: 5 } },
        { s: { r: 0, c: 6 }, e: { r: 0, c: 11 } },
        { s: { r: 0, c: 12 }, e: { r: 0, c: 17 } },
      ];

      ws['!cols'] = [
        { wch: 20 }, { wch: 35 }, { wch: 15 }, { wch: 35 }, { wch: 10 }, { wch: 10 },
        { wch: 15 }, { wch: 40 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 25 },
        { wch: 15 }, { wch: 40 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 25 }
      ];

      const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddr = XLSX.utils.encode_cell({ r, c });
          if (!ws[cellAddr]) ws[cellAddr] = { v: "" };
          ws[cellAddr].s = {
            border: {
              top: { style: "thin" }, bottom: { style: "thin" },
              left: { style: "thin" }, right: { style: "thin" }
            },
            alignment: { wrapText: true, vertical: "top", horizontal: "center" },
            font: { sz: 9, bold: r < 4 }
          };
        }
      }

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Реестр");

      // Добавление упрощенного реестра как второй лист
      if (allRegistryRows.length > 0) {
        const simplifiedHeaders = [
          "Регистрационный номер маршрута регулярных перевозок",
          "Порядковый номер маршрута регулярных перевозок",
          "Наименование маршрута регулярных перевозок в виде наименований начального остановочного пункта и конечного остановочного пункта по маршруту регулярных перевозок"
        ];
        
        const simplifiedData = allRegistryRows.map(row => [
          row['__col_0'],
          row['__col_1'],
          row['__col_2']
        ]);
        
        const wsSimplified = XLSX.utils.aoa_to_sheet([simplifiedHeaders, ...simplifiedData]);
        wsSimplified['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 80 }];
        XLSX.utils.book_append_sheet(wb, wsSimplified, "Отработанный реестр");
      }

      setResultWorkbook(wb);
      setResultSummary({ route: routeNumberFromGtfs, stopsCount: stopSequence.length, registryMatch: !!registryEntry });
    } catch (err: any) {
      setError(err.message || 'Ошибка генерации');
    } finally {
      setProcessing(false);
    }
  };

  const downloadDecoderResult = () => {
    if (decoderWorkbook) XLSX.writeFile(decoderWorkbook, `Расписание_остановок_${new Date().toLocaleDateString()}.xlsx`);
  };

  const downloadResult = () => {
    if (resultWorkbook && resultSummary) {
      XLSX.writeFile(resultWorkbook, `Итоговый_реестр_маршрут_${resultSummary.route}.xlsx`);
    }
  };

  const processRegistry = (type: MainFileType) => {
    setFiles(prev => ({
      ...prev,
      [type]: {
        ...prev[type],
        status: 'processed'
      }
    }));
  };

  const downloadProcessedRegistry = (type: MainFileType) => {
    const data = files[type].data;
    if (!data || data.length === 0) return;
    
    const headers = [
      "Регистрационный номер маршрута регулярных перевозок",
      "Порядковый номер маршрута регулярных перевозок",
      "Наименование маршрута регулярных перевозок в виде наименований начального остановочного пункта и конечного остановочного пункта по маршруту регулярных перевозок"
    ];
    
    const sheetData = data.map(row => [row['__col_0'], row['__col_1'], row['__col_2']]);
    
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sheetData]);
    ws['!cols'] = [{ wch: 30 }, { wch: 15 }, { wch: 80 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Отработанный реестр");
    
    const fileName = type === 'registry_mun' 
      ? 'Обработанный_реестр_муниципальный.xlsx' 
      : 'Обработанный_реестр_межмуниципальный.xlsx';
    XLSX.writeFile(wb, fileName);
  };

  const removeFile = (type: MainFileType) => {
    setFiles(prev => ({ ...prev, [type]: { file: null, data: null, status: 'empty' } }));
  };

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (password === 'trans2026') {
      setIsAuthenticated(true);
      localStorage.setItem('auth', 'true');
    } else {
      setAuthError(true);
      setTimeout(() => setAuthError(false), 2000);
    }
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-[#EEF2F6] font-sans flex flex-col items-center justify-center relative p-4">
        {/* Decorative background elements */}
        <div className="absolute inset-0 overflow-hidden pointer-events-none">
          <div className="absolute top-[-10%] sm:top-[-20%] left-[-10%] w-[300px] h-[300px] sm:w-[500px] sm:h-[500px] bg-[#00A478]/5 rounded-full blur-3xl"></div>
          <div className="absolute bottom-[-10%] sm:bottom-[-20%] right-[-10%] w-[300px] h-[300px] sm:w-[500px] sm:h-[500px] bg-[#D43A3A]/5 rounded-full blur-3xl"></div>
        </div>

        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-6 sm:p-10 rounded-[32px] sm:rounded-[40px] shadow-[0_20px_50px_rgba(0,0,0,0.05)] border border-white w-full max-w-md relative z-10"
        >
          <div className="flex flex-col items-center mb-8 sm:mb-10 text-center">
            <div className="w-14 sm:w-16 h-16 sm:h-20 bg-[#D43A3A] rounded-sm flex items-center justify-center p-1 sm:p-1.5 mb-4 sm:mb-6 shadow-md">
              <div className="w-8 sm:w-10 h-8 sm:h-10 bg-yellow-400 rounded-sm"></div>
            </div>
            <h1 className="text-[20px] sm:text-[24px] font-black text-[#1A237E] uppercase tracking-tighter leading-tight mb-2">
              АВТОРИЗАЦИЯ
            </h1>
            <p className="text-[11px] sm:text-[12px] font-bold text-[#A2AAB8] uppercase tracking-widest text-center">
              МИНИСТЕРСТВО ТРАНСПОРТА<br/>ТВЕРСКОЙ ОБЛАСТИ
            </p>
          </div>

          <form onSubmit={handleLogin} className="flex flex-col gap-4 sm:gap-6">
            <div className="relative">
              <div className="absolute inset-y-0 left-0 pl-4 sm:pl-5 flex items-center pointer-events-none">
                <Lock className="h-4 w-4 sm:h-5 sm:w-5 text-[#A2AAB8]" />
              </div>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="ПАРОЛЬ ДОСТУПА"
                className={cn(
                  "w-full pl-12 sm:pl-14 pr-4 py-4 sm:py-5 bg-[#F8FAFC] border rounded-xl text-[12px] sm:text-[14px] font-bold text-[#1A237E] placeholder-[#A2AAB8] transition-all outline-none uppercase tracking-widest",
                  authError ? "border-[#D43A3A] bg-[#FEF2F2] text-[#D43A3A]" : "border-[#DEE4ED] focus:border-[#00A478] focus:ring-1 focus:ring-[#00A478]"
                )}
                autoFocus
              />
              <AnimatePresence>
                {authError && (
                  <motion.div 
                    initial={{ opacity: 0, x: -10 }} 
                    animate={{ opacity: 1, x: 0 }} 
                    exit={{ opacity: 0 }}
                    className="absolute -bottom-6 left-0 text-[10px] font-bold text-[#D43A3A] uppercase tracking-wider"
                  >
                    Неверный пароль
                  </motion.div>
                )}
              </AnimatePresence>
            </div>

            <button
              type="submit"
              className="w-full py-4 sm:py-5 mt-2 rounded-xl text-[12px] sm:text-[14px] font-bold uppercase tracking-widest transition-all shadow-md bg-[#00A478] hover:bg-[#008B65] text-white active:scale-95 flex justify-center items-center gap-2 sm:gap-3"
            >
              <KeyRound className="w-4 h-4 sm:w-5 sm:h-5" />
              <span>ВОЙТИ В СИСТЕМУ</span>
            </button>
            <p className="text-center text-[9px] sm:text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-2 sm:mt-4">
              Рабочее место диспетчера
            </p>
          </form>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#EEF2F6] font-sans flex flex-col items-center relative pb-20">
      {/* Official Header */}
      <header className="w-full bg-[#1A237E] text-white shadow-lg sticky top-0 z-50">
        <div className="mx-auto px-10 h-20 flex justify-between items-center relative">
          <div className="flex items-center gap-6">
            <div className="flex items-center gap-4 cursor-pointer" onClick={() => setActiveTab('main')}>
              <div className="w-10 h-12 bg-[#D43A3A] rounded-sm flex items-center justify-center p-1">
                <div className="w-6 h-6 bg-yellow-400 rounded-sm"></div>
              </div>
              <div className="flex flex-col">
                <span className="text-[10px] font-medium leading-tight uppercase tracking-widest text-[#A2AAB8]">Министерство транспорта</span>
                <span className="text-[12px] font-bold uppercase tracking-widest text-white">Тверской области</span>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-6">
            {/* User Profile */}
            <div 
              className="bg-[#DFE4EC] text-[#1A237E] w-10 h-10 rounded-full flex items-center justify-center cursor-pointer hover:bg-white transition-all shadow-sm relative group"
              onClick={() => {
                setIsAuthenticated(false);
                localStorage.removeItem('auth');
              }}
            >
              <User className="w-5 h-5" />
              <div className="absolute top-12 right-0 bg-white shadow-lg border border-slate-200 rounded-lg px-4 py-2 text-[10px] uppercase tracking-widest font-bold opacity-0 group-hover:opacity-100 pointer-events-none transition-opacity whitespace-nowrap text-[#1A237E]">
                Выйти
              </div>
            </div>
          </div>
        </div>
      </header>



      <div className="w-full max-w-[1400px] px-6 md:px-12 py-4 md:py-6 flex-grow flex flex-col">
        <div className="flex flex-wrap gap-3 mb-6 justify-center">
          <button 
            translate="no"
            onClick={() => setActiveTab('decoder')}
            className={cn(
              "px-6 py-3 text-[12px] font-bold uppercase rounded-xl transition-all flex items-center justify-center gap-3",
              activeTab === 'decoder' 
                ? "bg-white text-[#1A237E] shadow-md border border-[#DEE4ED]" 
                : "bg-[#E9EEF4] text-[#1A237E] border border-[#DEE4ED] opacity-60 hover:opacity-100"
            )}
          >
            <Calendar className={cn("w-4 h-4", activeTab === 'decoder' ? "text-[#1A237E]" : "text-[#94A3B8]")} />
            <span className="text-center leading-tight">Формирование<br/>расписания</span>
          </button>
          <button 
            translate="no"
            onClick={() => setActiveTab('main')}
            className={cn(
              "px-6 py-3 text-[12px] font-bold uppercase rounded-xl transition-all flex items-center justify-center gap-3",
              activeTab === 'main'
                ? "bg-white text-[#1A237E] shadow-md border border-[#DEE4ED]"
                : "bg-[#E9EEF4] text-[#1A237E] border border-[#DEE4ED] opacity-60 hover:opacity-100"
            )}
          >
            <List className={cn("w-4 h-4", activeTab === 'main' ? "text-[#1A237E]" : "text-[#94A3B8]")} />
            <span className="text-center leading-tight">Формирование реестра<br/>расписаний</span>
          </button>
        </div>

        <AnimatePresence mode="wait">
          {activeTab === 'main' && (
            <motion.div 
              key="main"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="grid grid-cols-1 lg:grid-cols-12 gap-6"
            >
              {/* Left Content Area */}
              <div className="lg:col-span-7 flex flex-col">
                <div className="border-l-[4px] border-[#00A478] pl-4 mb-6">
                  <h2 className="text-[28px] font-black leading-tight text-[#1A237E] uppercase tracking-tighter mb-1">
                    ФОРМИРОВАНИЕ РЕЕСТРА
                  </h2>
                  <p className="text-[12px] font-semibold text-[#00A478] uppercase tracking-wide leading-relaxed">
                    АВТОМАТИЗИРОВАННАЯ ИНФОРМАЦИОННАЯ СИСТЕМА<br/>ИНТЕГРАЦИИ ДАННЫХ
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-2">
                  {Object.entries(FILE_CONFIG).map(([key, config]) => (
                    <div key={key} className="flex flex-col gap-2">

                      <FileCard
                        type={key as any}
                        config={config}
                        state={files[key as MainFileType]}
                        onUpload={(file) => handleFileChange(key as MainFileType, file)}
                        onRemove={() => removeFile(key as MainFileType)}
                        accept=".csv, .xlsx, .xls"
                      />
                      {['success', 'processed'].includes(files[key as MainFileType].status) && (
                        <div className="flex gap-2 w-full mt-2">
                          <button 
                            translate="no"
                            onClick={() => processRegistry(key as MainFileType)}
                            disabled={files[key as MainFileType].status === 'processed'}
                            className={cn(
                              "flex-1 py-3 text-[10px] font-black uppercase tracking-widest rounded shadow-md transition-all flex items-center justify-center gap-2",
                              files[key as MainFileType].status === 'processed' 
                                ? "bg-emerald-50 text-emerald-600 border border-emerald-200 cursor-not-allowed" 
                                : "bg-[#00A478] text-white hover:bg-[#008B65] active:scale-95"
                            )}
                          >
                            <span>{files[key as MainFileType].status === 'processed' ? 'ГОТОВО' : 'ОБРАБОТАТЬ'}</span>
                          </button>
                          {files[key as MainFileType].status === 'processed' && (
                            <button 
                              onClick={() => downloadProcessedRegistry(key as MainFileType)}
                              className="flex-1 py-3 bg-[#E31E24] text-white hover:bg-[#c0181d] text-[10px] font-black uppercase tracking-widest rounded shadow-md transition-all flex items-center justify-center gap-2 active:scale-95"
                            >
                              <Download className="w-4 h-4" />
                              <span>Скачать .xlsx</span>
                            </button>
                          )}
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                <div className="max-w-md mx-auto mt-4 mb-2">
                  <FileCard 
                    type="accumulator"
                    config={{ label: "Накопительный отчет (опционально)", icon: FileSpreadsheet, description: "Сюда будут добавляться новые маршруты" }}
                    state={accumulatorReportFile}
                    onUpload={handleAccumulatorReportFile}
                    onRemove={() => { setAccumulatorReportFile({ file: null, data: null, status: 'empty' }); }}
                    accept=".csv, .xlsx, .xls"
                  />
                  <p className="mt-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest text-center leading-relaxed">
                    Если загрузить ранее скачанный Итоговый отчет, данные нового маршрута добавятся в конец файла
                  </p>
                </div>
              </div>

              {/* Right Interactive Area */}
              <div className="lg:col-span-5 flex flex-col h-full pl-0 lg:pl-4">
                <div className="bg-[#F8FAFC] rounded-[24px] shadow-[0_20px_50px_rgba(0,0,0,0.02)] border border-white flex-grow flex flex-col items-center justify-center text-center relative p-6 overflow-hidden w-full backdrop-blur-md">
                    <div className="relative z-10 w-full flex flex-col items-center">
                      <div className="w-20 h-20 bg-[#EEF2F6] rounded-full flex items-center justify-center mx-auto mb-4 shadow-inner">
                        <Bus className="w-10 h-10 text-[#94A3B8]" />
                      </div>
                      
                      <div className="mb-6 w-full min-h-[100px] flex flex-col items-center justify-center">
                        {error ? (
                          <div className="space-y-4">
                            <AlertCircle className="w-14 h-14 text-[#D43A3A] mx-auto" />
                            <h4 className="text-[16px] font-bold text-[#D43A3A] uppercase tracking-widest">Ошибка</h4>
                            <p className="text-[11px] font-semibold text-slate-500 uppercase leading-relaxed max-w-[200px] mx-auto">
                              {error}
                            </p>
                          </div>
                        ) : processing ? (
                          <div className="space-y-4">
                            <motion.div animate={{ rotate: 360 }} transition={{ repeat: Infinity, duration: 2, ease: "linear" }}>
                              <History className="w-12 h-12 text-[#00A478] mx-auto" />
                            </motion.div>
                            <h4 className="text-[16px] font-bold text-[#1A237E] uppercase tracking-widest">Анализ данных...</h4>
                          </div>
                        ) : resultWorkbook ? (
                          <div className="space-y-4">
                            <motion.div initial={{ scale: 0 }} animate={{ scale: 1 }} className="flex justify-center -mt-8 mb-4">
                              <CheckCircle2 className="w-14 h-14 text-[#00A478] bg-white rounded-full p-1" />
                            </motion.div>
                            <h4 className="text-[24px] font-black text-[#1A237E] uppercase tracking-normal">РЕЗУЛЬТАТ ГОТОВ</h4>
                            <div className="space-y-1">
                              <p className="text-[12px] font-bold text-[#00A478] uppercase tracking-wider">
                                {resultSummary?.route} МАРШРУТ ({resultSummary?.stopsCount} ОСТ.)
                              </p>
                              <p className={cn("text-[10px] font-bold uppercase tracking-wider mt-2", resultSummary?.registryMatch ? "text-[#00A478]" : "text-amber-500")}>
                                {resultSummary?.registryMatch ? "ДАННЫЕ ИЗ РЕЕСТРА ДОБАВЛЕНЫ" : "ДАННЫЕ В РЕЕСТРЕ НЕ НАЙДЕНЫ (СТОЛБЦЫ БУДУТ ПУСТЫМИ)"}
                              </p>
                            </div>
                          </div>
                        ) : (
                          <div className="space-y-4 opacity-50">
                            <h4 className="text-xl font-bold text-[#1A237E] uppercase tracking-wider leading-tight">Ожидание данных</h4>
                            <p className="text-[12px] font-medium text-slate-400 uppercase tracking-wide leading-relaxed max-w-[200px] mx-auto">
                              Загрузите файлы для результата
                            </p>
                          </div>
                        )}
                      </div>

                      <div className="flex flex-col w-full gap-6 px-4">
                        <button
                          translate="no"
                          onClick={resultWorkbook ? downloadResult : generateReport}
                          disabled={!resultWorkbook && (processing || (!winterDecoderFile.data && !summerDecoderFile.data))}
                          className={cn(
                            "w-full py-5 rounded-full text-[14px] font-bold uppercase tracking-widest transition-all flex items-center justify-center gap-3",
                            !resultWorkbook && (processing || (!winterDecoderFile.data && !summerDecoderFile.data)) && "bg-[#DEE4ED] text-[#94A3B8] cursor-not-allowed",
                            ((winterDecoderFile.data && winterDecoderFile.data.length > 0) || (summerDecoderFile.data && summerDecoderFile.data.length > 0)) && !processing && !resultWorkbook && "bg-[#00A478] text-white hover:bg-[#008B65] shadow-xl active:scale-95 cursor-pointer",
                            resultWorkbook && "bg-[#00A478]/90 backdrop-blur-md border border-white/30 text-white hover:bg-[#008B65] shadow-[0_8px_32px_rgba(0,164,120,0.4)] active:scale-95"
                          )}
                        >
                          {processing ? <History className="w-5 h-5 animate-spin" /> : <Download className="w-5 h-5" />}
                          <span>{processing ? 'ВЫПОЛНЕНИЕ...' : resultWorkbook ? 'СКАЧАТЬ ОТЧЕТ' : 'СОЗДАТЬ ОТЧЕТ'}</span>
                        </button>
                        
                        {(resultWorkbook || processing) ? (
                            <button 
                              onClick={() => { setResultWorkbook(null); setResultSummary(null); }}
                              className="text-[12px] font-bold text-[#D43A3A] uppercase tracking-widest hover:underline mx-auto"
                            >
                              СБРОСИТЬ СЕССИЮ
                            </button>
                        ) : (
                           <div className="h-4"></div>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              </motion.div>
          )}
          {activeTab === 'decoder' && (
            <motion.div 
              key="decoder"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="max-w-3xl mx-auto"
            >
              <div className="bg-white p-8 rounded-[24px] border border-[#DEE4ED] shadow-sm text-center relative overflow-hidden">
                <div className="absolute top-0 left-0 w-full h-[6px] bg-[#00A478]"></div>
                
                <div className="w-16 h-16 bg-[#F0F4F8] rounded-2xl mx-auto mb-6 flex items-center justify-center">
                  <FileSpreadsheet className="w-8 h-8 text-[#00A478]" />
                </div>
                
                <h2 className="text-2xl font-black uppercase tracking-tighter text-[#1A237E] mb-2">ФОРМИРОВАНИЕ РАСПИСАНИЯ</h2>
                <p className="text-[#94A3B8] text-[11px] uppercase tracking-widest font-bold mb-6">Преобразование числовых последовательностей GTFS</p>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 max-w-2xl mx-auto mb-8">
                  <FileCard 
                    type="winter_gtfs"
                    config={{ label: "Входящий массив (Зима)", icon: FileText, description: "TXT / CSV" }}
                    state={winterDecoderFile}
                    onUpload={handleWinterDecoderFile}
                    onRemove={() => {
                      setWinterDecoderFile({ file: null, data: null, status: 'empty' });
                      setDecoderWorkbook(null);
                    }}
                  />
                  <FileCard 
                    type="summer_gtfs"
                    config={{ label: "Входящий массив (Лето)", icon: FileText, description: "TXT / CSV" }}
                    state={summerDecoderFile}
                    onUpload={handleSummerDecoderFile}
                    onRemove={() => {
                      setSummerDecoderFile({ file: null, data: null, status: 'empty' });
                      setDecoderWorkbook(null);
                    }}
                  />
                </div>

                <div className="flex flex-col gap-4 max-w-xs mx-auto">
                  {error && (
                    <div className="bg-red-50 border border-red-100 p-4 rounded-xl mb-4">
                      <p className="text-[10px] font-bold text-red-600 uppercase tracking-widest leading-relaxed">
                        {error}
                      </p>
                    </div>
                  )}
                  <button
                    onClick={runDecoder}
                    disabled={winterDecoderFile.status !== 'success' && summerDecoderFile.status !== 'success'}
                    className={cn(
                      "w-full py-5 rounded-full text-[14px] font-bold uppercase tracking-widest transition-all flex items-center justify-center gap-3",
                      winterDecoderFile.status !== 'success' && summerDecoderFile.status !== 'success' ? "bg-[#DEE4ED] text-[#94A3B8] cursor-not-allowed" : "bg-[#00A478] text-white hover:bg-[#008B65] shadow-lg active:scale-95",
                      decoderWorkbook && "opacity-50"
                    )}
                  >
                    <Download className="w-5 h-5" />
                    <span>Сформировать</span>
                  </button>

                  {decoderWorkbook && (
                    <button
                      onClick={downloadDecoderResult}
                      className="w-full py-5 rounded-full text-[14px] font-bold uppercase tracking-widest bg-[#00A478]/90 backdrop-blur-md border border-white/30 hover:bg-[#008B65] text-white transition-all shadow-[0_8px_32px_rgba(0,164,120,0.4)] flex items-center justify-center gap-3 active:scale-95"
                    >
                      <Download className="w-5 h-5" />
                      <span>Скачать</span>
                    </button>
                  )}
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>

    </div>
  );
}
