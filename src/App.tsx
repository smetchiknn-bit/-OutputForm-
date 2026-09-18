import { useState, useRef } from 'react';
import { processWorkbookWith1C, exportToExcel, getRowType, getRowStyle } from './utils/excelProcessor';

export default function App() {
  const [result, setResult] = useState<any>(null);
  const [logs, setLogs] = useState<any[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [activeSheet, setActiveSheet] = useState<number>(0);
  const [pendingFile, setPendingFile] = useState<ArrayBuffer | null>(null);
  const [showVersionHistory, setShowVersionHistory] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const versionHistory = [
    {
      version: "1.0.0",
      date: "2024-01-15",
      changes: [
        "Первая стабильная версия",
        "Обработка листа Свод с 19 колонками (A-S)",
        "Нумерация иерархии (О, К, С, У, Э, Л1-Л3, ГР, КЕР, ТМЦ)",
        "Формулы для расчёта СМР и ТМЦ",
        "Форматирование ячеек по типам строк",
        "Обработка листов СМР уник и ТМЦ уник",
        "Экспорт в Excel с сохранением форматирования"
      ]
    }
  ];

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      const arrayBuffer = e.target?.result as ArrayBuffer;
      setPendingFile(arrayBuffer);
      
      const { result: res, logs: lgs } = processWorkbookWith1C(arrayBuffer, false);
      setResult(res);
      setLogs(lgs);
      setActiveSheet(0);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleExport = () => {
    if (result?.workbook) {
      exportToExcel(result.workbook, fileName.replace('.xlsx', '_обработанный.xlsx'));
    }
  };

  const handleReset = () => {
    setResult(null);
    setLogs([]);
    setFileName('');
    setActiveSheet(0);
    setPendingFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="min-h-screen">
      <div className="floating-symbol">Σ</div>
      <div className="floating-symbol">₽</div>
      <div className="floating-symbol">=</div>

      <header className="bg-[#f0fdf4] text-black shadow-lg" style={{ boxShadow: 'var(--shadow-hard-green)' }}>
        <div className="max-w-7xl mx-auto px-4 py-6">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-4">
              <img src="/logoStiker.png" alt="Стикер" className="w-12 h-12 rounded-lg" style={{ boxShadow: 'var(--shadow-hard)' }} />
              <div>
                <h1 className="text-2xl font-bold text-black">Выходная форма</h1>
                <p className="text-sm text-black/80">обработка репорт Стикер 2.0 · нумерация · формулы</p>
              </div>
            </div>
            <img src="/logoXLSX.png" alt="Excel" className="w-12 h-12 rounded-lg" style={{ boxShadow: 'var(--shadow-hard)' }} />
          </div>
          
          {fileName && (
            <div className="mt-4 flex items-center justify-between bg-white/50 rounded-lg px-4 py-2">
              <span className="font-mono text-sm text-black">{fileName}</span>
              <div className="flex items-center gap-3">
                {result && (
                  <button
                    onClick={handleExport}
                    className="flex items-center gap-2 bg-[#16a34a] text-white px-4 py-2 rounded-lg font-semibold hover:bg-[#14532d] transition-colors"
                    style={{ boxShadow: 'var(--shadow-hard)' }}
                  >
                    <img src="/logoXLSX.png" alt="Excel" className="w-5 h-5" />
                    <span>Скачать</span>
                  </button>
                )}
                <button
                  onClick={handleReset}
                  className="text-sm text-black hover:text-red-600 transition-colors"
                >
                  Сброс
                </button>
              </div>
            </div>
          )}
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
        {!result ? (
          <div className="max-w-2xl mx-auto">
            <div className="bg-white rounded-xl p-12 border-2 border-dashed border-[#86efac] hover:border-[#16a34a] transition-colors" style={{ boxShadow: 'var(--shadow-hard)' }}>
              <div className="text-center">
                <div className="w-20 h-20 mx-auto mb-6 bg-[#16a34a] rounded-full flex items-center justify-center" style={{ boxShadow: 'var(--shadow-hard-green)' }}>
                  <svg className="w-10 h-10 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                </div>
                <h2 className="text-2xl font-bold text-[#14532d] mb-2">
                  Перетащите репорт Стикер 2.0 сюда
                </h2>
                <p className="text-[#14532d]/70 mb-6">или нажмите, чтобы выбрать файл · .xlsx</p>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileSelect}
                  className="hidden"
                />
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="bg-[#16a34a] hover:bg-[#14532d] text-white px-8 py-3 rounded-lg font-semibold transition-colors"
                  style={{ boxShadow: 'var(--shadow-hard-green)' }}
                >
                  Выбрать файл
                </button>
                <p className="text-xs text-[#14532d]/60 mt-6">
                  💡 Файл обрабатывается локально в браузере и никуда не отправляется
                </p>
              </div>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6" style={{ boxShadow: 'var(--shadow-hard)' }}>
              <h2 className="text-lg font-bold text-[#14532d] mb-4">Журнал обработки</h2>
              <div className="space-y-2">
                {logs.map((log, idx) => (
                  <div key={idx} className="flex items-center gap-3 text-sm">
                    <span className={`w-2 h-2 rounded-full ${
                      log.status === 'success' ? 'bg-[#16a34a]' :
                      log.status === 'error' ? 'bg-red-500' : 'bg-yellow-500'
                    }`}></span>
                    <span className="font-semibold text-[#14532d]">{log.step}:</span>
                    <span className="text-[#14532d]/70">{log.message}</span>
                  </div>
                ))}
              </div>
            </div>

            <div className="bg-white rounded-xl overflow-hidden" style={{ boxShadow: 'var(--shadow-hard)' }}>
              <div className="border-b border-[#86efac]">
                <div className="flex overflow-x-auto">
                  {result.sheets?.map((sheet: any, idx: number) => (
                    <button
                      key={idx}
                      onClick={() => setActiveSheet(idx)}
                      className={`px-6 py-3 font-semibold whitespace-nowrap transition-colors ${
                        activeSheet === idx
                          ? 'bg-[#16a34a] text-white'
                          : 'bg-[#f0fdf4] text-[#14532d] hover:bg-[#dcfce7]'
                      }`}
                    >
                      {sheet.name}
                    </button>
                  ))}
                </div>
              </div>

              <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                <table className="w-full text-sm">
                  <thead className="bg-[#dcfce7] sticky top-0">
                    <tr>
                      {(() => {
                        const numCols = activeSheet === 0 ? 10 : (result.sheets?.[activeSheet]?.data[0]?.length || 0);
                        return Array.from({ length: numCols }, (_, colIdx) => {
                          const cell = result.sheets?.[activeSheet]?.data[0]?.[colIdx];
                          return (
                            <th key={colIdx} className="px-3 py-2 text-left font-semibold text-[#14532d] border-b-2 border-[#16a34a] whitespace-pre-line">
                              {cell || ''}
                            </th>
                          );
                        });
                      })()}
                    </tr>
                  </thead>
                  <tbody>
                    {result.sheets?.[activeSheet]?.data.slice(1).map((row: any[], rowIdx: number) => {
                      const type = activeSheet === 0 ? getRowType(result.sheets[0].data, rowIdx + 1) : '';
                      const numCols = activeSheet === 0 ? 10 : row.length;
                      
                      return (
                        <tr key={rowIdx} className={activeSheet === 0 ? getRowStyle(type) : 'hover:bg-[#f0fdf4]'}>
                          {Array.from({ length: numCols }, (_, colIdx) => {
                            const cell = row[colIdx];
                            let cellStyle = '';
                            
                            if (activeSheet === 0) {
                              cellStyle = getRowStyle(type, colIdx);
                            }
                            
                            return (
                              <td
                                key={colIdx}
                                className={`px-3 py-2 border-b border-[#86efac] ${cellStyle} text-[#14532d]`}
                              >
                                {cell !== null && cell !== undefined ? String(cell) : ''}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="flex justify-center gap-4">
              <button
                onClick={handleExport}
                className="bg-[#16a34a] hover:bg-[#14532d] text-white px-12 py-4 rounded-xl font-bold text-lg transition-all transform hover:scale-105 flex items-center gap-3"
                style={{ boxShadow: 'var(--shadow-hard-green)' }}
              >
                <img src="/logoXLSX.png" alt="Excel" className="w-8 h-8" />
                <span>Скачать Excel</span>
              </button>
            </div>
          </div>
        )}
      </main>

      {result && (
        <button
          onClick={handleExport}
          className="fixed bottom-8 right-8 bg-[#16a34a] hover:bg-[#14532d] text-white px-6 py-4 rounded-full font-bold transition-all transform hover:scale-110 flex items-center gap-3 border-4 border-white z-40"
          style={{ boxShadow: 'var(--shadow-hard-green)' }}
          title="Скачать Excel файл"
        >
          <img src="/logoXLSX.png" alt="Excel" className="w-8 h-8" />
          <span className="hidden sm:inline">Скачать</span>
        </button>
      )}

      <footer className="mt-12 py-6 text-center text-sm text-[#14532d]/60">
        <div className="max-w-7xl mx-auto px-4">
          <p className="mb-2">Выходная форма — обработка репорт Стикер 2.0 · нумерация · формулы</p>
          <button
            onClick={() => setShowVersionHistory(true)}
            className="text-xs text-[#14532d]/40 hover:text-[#16a34a] transition-colors underline"
          >
            Версия 1.0.0 от 15.01.2024
          </button>
        </div>
      </footer>

      {showVersionHistory && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-2xl w-full max-h-[80vh] overflow-hidden" style={{ boxShadow: 'var(--shadow-hard-green)' }}>
            <div className="bg-gradient-to-r from-[#14532d] to-[#16a34a] text-white px-6 py-4 flex items-center justify-between">
              <h2 className="text-xl font-bold">История стабильных версий</h2>
              <button
                onClick={() => setShowVersionHistory(false)}
                className="text-white hover:text-red-300 transition-colors text-2xl"
              >
                ×
              </button>
            </div>
            <div className="p-6 overflow-y-auto max-h-[60vh]">
              {versionHistory.map((version, idx) => (
                <div key={idx} className="mb-6 last:mb-0">
                  <div className="flex items-center gap-3 mb-3">
                    <span className="bg-[#16a34a] text-white px-3 py-1 rounded-lg font-bold text-sm">
                      v{version.version}
                    </span>
                    <span className="text-[#14532d]/60 text-sm">
                      {new Date(version.date).toLocaleDateString('ru-RU', {
                        year: 'numeric',
                        month: 'long',
                        day: 'numeric'
                      })}
                    </span>
                  </div>
                  <ul className="space-y-2">
                    {version.changes.map((change, changeIdx) => (
                      <li key={changeIdx} className="flex items-start gap-2 text-[#14532d]">
                        <span className="text-[#16a34a] mt-1">✓</span>
                        <span>{change}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
            <div className="bg-[#f0fdf4] px-6 py-4 border-t border-[#86efac]">
              <p className="text-xs text-[#14532d]/60">
                💡 Все файлы обрабатываются локально в вашем браузере и не отправляются на сервер
              </p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
