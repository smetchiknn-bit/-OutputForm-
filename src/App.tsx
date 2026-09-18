import { useState, useRef } from 'react';
import { processWorkbookWith1C, exportToExcel, getRowType, getRowStyle } from './utils/excelProcessor';

export default function App() {
  const [result, setResult] = useState<any>(null);
  const [logs, setLogs] = useState<any[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [activeSheet, setActiveSheet] = useState<number>(0);
  // const [showDialog1C, setShowDialog1C] = useState<boolean>(false); // Временно отключено
  const [pendingFile, setPendingFile] = useState<ArrayBuffer | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      const arrayBuffer = e.target?.result as ArrayBuffer;
      setPendingFile(arrayBuffer);
      
      // Сначала обрабатываем без листа 1С
      const { result: res, logs: lgs } = processWorkbookWith1C(arrayBuffer, false);
      setResult(res);
      setLogs(lgs);
      setActiveSheet(0);
      
      // Показываем диалог о добавлении листа 1С
      // setShowDialog1C(true); // Временно отключено
    };
    reader.readAsArrayBuffer(file);
  };

  // Временно отключено - создание листа 1С
  // const handleAddSheet1C = (add: boolean) => {
  //   if (add && pendingFile) {
  //     // Переобрабатываем с листом 1С
  //     const { result: res, logs: lgs } = processWorkbookWith1C(pendingFile, true);
  //     setResult(res);
  //     setLogs(lgs);
  //   }
  //   setShowDialog1C(false);
  // };

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
    // setShowDialog1C(false); // Временно отключено
    setPendingFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">
      {/* Header */}
      <header className="bg-gradient-to-r from-[#123f28] to-[#1e7145] text-white shadow-lg">
        <div className="max-w-7xl mx-auto px-4 py-6">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-4">
              <img src="/logoStiker.png" alt="Стикер" className="w-12 h-12 rounded-lg shadow-md" />
              <div>
                <h1 className="text-2xl font-bold">Выходная форма</h1>
                <p className="text-sm text-white/80">обработка репорт Стикер 2.0 · нумерация · формулы · лист 1С</p>
              </div>
            </div>
            <img src="/logoXLSX.png" alt="Excel" className="w-12 h-12 rounded-lg shadow-md" />
          </div>
          
          {fileName && (
            <div className="mt-4 flex items-center justify-between bg-white/10 rounded-lg px-4 py-2">
              <span className="font-mono text-sm">{fileName}</span>
              <div className="flex items-center gap-3">
                {result && (
                  <button
                    onClick={handleExport}
                    className="flex items-center gap-2 bg-white text-[#1e7145] px-4 py-2 rounded-lg font-semibold shadow-md hover:bg-slate-100 transition-colors"
                  >
                    <img src="/logoXLSX.png" alt="Excel" className="w-5 h-5" />
                    <span>Скачать</span>
                  </button>
                )}
                <button
                  onClick={handleReset}
                  className="text-sm hover:text-red-300 transition-colors"
                >
                  Сброс
                </button>
              </div>
            </div>
          )}
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 py-8">
        {!result ? (
          // Drop Zone
          <div className="max-w-2xl mx-auto">
            <div className="bg-white rounded-xl shadow-xl p-12 border-2 border-dashed border-slate-300 hover:border-[#1e7145] transition-colors">
              <div className="text-center">
                <div className="w-20 h-20 mx-auto mb-6 bg-[#1e7145] rounded-full flex items-center justify-center shadow-lg">
                  <svg className="w-10 h-10 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                </div>
                <h2 className="text-2xl font-bold text-slate-800 mb-2">
                  Перетащите репорт Стикер 2.0 сюда
                </h2>
                <p className="text-slate-600 mb-6">или нажмите, чтобы выбрать файл · .xlsx</p>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileSelect}
                  className="hidden"
                />
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="bg-[#1e7145] hover:bg-[#123f28] text-white px-8 py-3 rounded-lg font-semibold shadow-md transition-colors"
                >
                  Выбрать файл
                </button>
                <p className="text-xs text-slate-500 mt-6">
                  Файл обрабатывается локально в браузере и никуда не отправляется
                </p>
              </div>
            </div>
          </div>
        ) : (
          // Results
          <div className="space-y-6">
            {/* Logs */}
            <div className="bg-white rounded-xl shadow-lg p-6">
              <h2 className="text-lg font-bold text-slate-800 mb-4">Журнал обработки</h2>
              <div className="space-y-2">
                {logs.map((log, idx) => (
                  <div key={idx} className="flex items-center gap-3 text-sm">
                    <span className={`w-2 h-2 rounded-full ${
                      log.status === 'success' ? 'bg-green-500' :
                      log.status === 'error' ? 'bg-red-500' : 'bg-yellow-500'
                    }`}></span>
                    <span className="font-semibold text-slate-700">{log.step}:</span>
                    <span className="text-slate-600">{log.message}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Sheet Tabs */}
            <div className="bg-white rounded-xl shadow-lg overflow-hidden">
              <div className="border-b border-slate-200">
                <div className="flex overflow-x-auto">
                  {result.sheets?.map((sheet: any, idx: number) => (
                    <button
                      key={idx}
                      onClick={() => setActiveSheet(idx)}
                      className={`px-6 py-3 font-semibold whitespace-nowrap transition-colors ${
                        activeSheet === idx
                          ? 'bg-[#1e7145] text-white'
                          : 'bg-slate-50 text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      {sheet.name}
                    </button>
                  ))}
                </div>
              </div>

              {/* Table */}
              <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                <table className="w-full text-sm">
                  <thead className="bg-slate-100 sticky top-0">
                    <tr>
                      {(() => {
                        // Для листа Свод показываем только первые 10 колонок (A-J)
                        const numCols = activeSheet === 0 ? 10 : (result.sheets?.[activeSheet]?.data[0]?.length || 0);
                        return Array.from({ length: numCols }, (_, colIdx) => {
                          const cell = result.sheets?.[activeSheet]?.data[0]?.[colIdx];
                          return (
                            <th key={colIdx} className="px-3 py-2 text-left font-semibold text-slate-700 border-b border-slate-300 whitespace-pre-line">
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
                      // Для листа Свод показываем все 10 колонок (A-J)
                      const numCols = activeSheet === 0 ? 10 : row.length;
                      
                      return (
                        <tr key={rowIdx} className={activeSheet === 0 ? getRowStyle(type) : ''}>
                          {Array.from({ length: numCols }, (_, colIdx) => {
                            const cell = row[colIdx];
                            let cellStyle = '';
                            
                            if (activeSheet === 0) {
                              cellStyle = getRowStyle(type, colIdx);
                            }
                            
                            return (
                              <td
                                key={colIdx}
                                className={`px-3 py-2 border-b border-slate-200 ${cellStyle} text-slate-800`}
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

            {/* Export Button */}
            <div className="flex justify-center gap-4">
              <button
                onClick={handleExport}
                className="bg-gradient-to-r from-[#1e7145] to-[#2d8a5e] hover:from-[#123f28] hover:to-[#1e7145] text-white px-12 py-4 rounded-xl font-bold text-lg shadow-2xl transition-all transform hover:scale-105 flex items-center gap-3 border-2 border-white"
              >
                <img src="/logoXLSX.png" alt="Excel" className="w-8 h-8" />
                <span>Скачать Excel</span>
              </button>
            </div>
          </div>
        )}
      </main>

      {/* Dialog 1C - Временно отключено */}
      {/* {showDialog1C && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white rounded-xl shadow-2xl p-8 max-w-md">
            <div className="flex items-center gap-4 mb-6">
              <img src="/logo1C.png" alt="1С" className="w-12 h-12" />
              <h3 className="text-xl font-bold text-slate-800">Добавить лист «1С»?</h3>
            </div>
            <p className="text-slate-600 mb-6">
              Основной файл обработан. Хотите добавить дополнительный лист «1С» с форматированием для выгрузки в 1С?
            </p>
            <div className="flex gap-3">
              <button
                onClick={() => handleAddSheet1C(true)}
                className="flex-1 bg-[#1e7145] hover:bg-[#123f28] text-white px-4 py-2 rounded-lg font-semibold shadow-md transition-colors"
              >
                Да, добавить
              </button>
              <button
                onClick={() => handleAddSheet1C(false)}
                className="flex-1 bg-slate-200 hover:bg-slate-300 text-slate-700 px-4 py-2 rounded-lg font-semibold transition-colors"
              >
                Нет, пропустить
              </button>
            </div>
          </div>
        </div>
      )} */}

      {/* Floating Download Button */}
      {result && (
        <button
          onClick={handleExport}
          className="fixed bottom-8 right-8 bg-gradient-to-r from-[#1e7145] to-[#2d8a5e] hover:from-[#123f28] hover:to-[#1e7145] text-white px-6 py-4 rounded-full font-bold shadow-2xl transition-all transform hover:scale-110 flex items-center gap-3 border-4 border-white z-40"
          title="Скачать Excel файл"
        >
          <img src="/logoXLSX.png" alt="Excel" className="w-8 h-8" />
          <span className="hidden sm:inline">Скачать</span>
        </button>
      )}

      {/* Footer */}
      <footer className="mt-12 py-6 text-center text-sm text-slate-500">
        Выходная форма — обработка репорт Стикер 2.0 · нумерация · формулы · лист 1С
      </footer>
    </div>
  );
}
