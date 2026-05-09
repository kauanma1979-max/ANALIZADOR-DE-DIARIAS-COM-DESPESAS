/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  PointElement,
  LineElement,
  ArcElement,
} from 'chart.js';
import { Bar, Line, Doughnut } from 'react-chartjs-2';
import { 
  FileSpreadsheet, 
  Upload, 
  Download, 
  Trash2, 
  RefreshCw, 
  PieChart, 
  BarChart3, 
  Table as TableIcon, 
  Plus,
  Coins,
  Search,
  Calendar,
  Layers,
  Save,
  FolderOpen,
  ExternalLink,
  MessageCircle,
  Edit2
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import { cn } from './lib/utils';
import { DiariaRecord, Expense, TabType, MESES_NUMERO, MONTH_ORDER } from './types';

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  PointElement,
  LineElement,
  ArcElement,
  Title,
  Tooltip,
  Legend
);

export default function App() {
  const [allData, setAllData] = useState<DiariaRecord[]>([]);
  const [expenses, setExpenses] = useState<Expense[]>([]);
  const [activeTab, setActiveTab] = useState<TabType>('dados');
  const [isLoading, setIsLoading] = useState(false);
  const [isSyncingExpenses, setIsSyncingExpenses] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState<string | null>(null);
  const [mesFilter, setMesFilter] = useState('all');
  const [anoFilter, setAnoFilter] = useState('all');
  const [searchFilter, setSearchFilter] = useState('');
  
  const [cityFilterExpenses, setCityFilterExpenses] = useState('all');
  const [showOnlyAccommodation, setShowOnlyAccommodation] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  // Sync expenses from Google Sheets
  const syncExpenses = async () => {
    setIsSyncingExpenses(true);
    try {
      const PROXY_URL = '/api/sync-expenses';
      const DIRECT_URL = 'https://docs.google.com/spreadsheets/d/1lyXkSmeiyyODZbng6GtXTSwNR-XY2KWRLKWqppLef1k/export?format=csv&gid=980751451';
      
      let csvText = '';
      let response;

      // Try proxy first
      try {
        response = await fetch(PROXY_URL);
        if (response.ok) {
          csvText = await response.text();
        } else if (response.status === 404) {
          console.warn('Proxy returned 404, trying direct fetch...');
        } else {
          const err = await response.json().catch(() => ({}));
          throw new Error(err.error || `Erro ${response.status}`);
        }
      } catch (proxyErr) {
        console.error('Proxy failed:', proxyErr);
      }

      // Fallback to direct fetch if proxy failed or returned 404
      if (!csvText) {
        console.log('Attempting direct fetch...');
        response = await fetch(DIRECT_URL);
        if (!response.ok) throw new Error('Não foi possível carregar os dados da planilha (Proxy e Direto falharam)');
        csvText = await response.text();
      }
      
      if (!csvText || csvText.length < 50) {
        throw new Error('Os dados recebidos da planilha parecem estar vazios ou inválidos.');
      }

      const workbook = XLSX.read(csvText, { type: 'string' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      
      // Busca a linha de cabeçalho de forma robusta para pular linhas vazias no topo
      const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      const headerRowIndex = rawRows.findIndex(row => 
        row.some(cell => cell && String(cell).toUpperCase().includes('DATA'))
      );

      if (headerRowIndex === -1) {
        throw new Error('Não foi possível encontrar a coluna "DATA" na planilha.');
      }

      const headers = rawRows[headerRowIndex].map(h => String(h || '').trim().toUpperCase());
      const rawData = rawRows.slice(headerRowIndex + 1).map(row => {
        const obj: any = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

      const mappedExpenses: Expense[] = rawData.map((normalizedRow, idx) => {
        let dateVal = normalizedRow['DATA'] || normalizedRow['DATE'] || '';
        let formattedDate = toISODate(dateVal);

        const rawValor = normalizedRow['VALOR'] || '0';
        let numericValue = 0;
        if (typeof rawValor === 'number') {
          numericValue = rawValor;
        } else {
          const valStr = rawValor.toString()
            .replace('R$', '')
            .replace(/\./g, '')
            .replace(',', '.')
            .trim();
          numericValue = parseFloat(valStr) || 0;
        }

        return {
          id: idx,
          date: formattedDate,
          value: numericValue,
          type: String(normalizedRow['TIPO DE DESPESA'] || 'Outros'),
          description: String(normalizedRow['DESCRIÇÃO'] || ''),
          city: String(normalizedRow['CIDADE'] || ''),
          link: String(normalizedRow['LINK'] || ''),
          whatsapp: String(normalizedRow['WHATSAPP'] || '')
        };
      }).filter(exp => exp.date && exp.value > 0); // Remove itens sem data ou valor zero

      setExpenses(mappedExpenses);
      setLastSyncTime(new Date().toLocaleTimeString('pt-BR'));
      localStorage.setItem('diarias_expenses', JSON.stringify(mappedExpenses));
      localStorage.setItem('diarias_last_sync', new Date().toLocaleTimeString('pt-BR'));
    } catch (error) {
      console.error('Erro ao sincronizar despesas:', error);
      alert('Erro ao sincronizar com a planilha do Google. Verifique sua conexão.');
    } finally {
      setIsSyncingExpenses(false);
    }
  };

  // Sync on tab change
  useEffect(() => {
    if (activeTab === 'despesas') {
      syncExpenses();
    }
  }, [activeTab]);

  // Load cache on mount
  useEffect(() => {
    const saved = localStorage.getItem('diarias_expenses');
    const syncTime = localStorage.getItem('diarias_last_sync');
    if (saved) {
      try {
        setExpenses(JSON.parse(saved));
        setLastSyncTime(syncTime);
      } catch (e) {
        console.error('Error loading expenses from cache', e);
      }
    }
  }, []);

  const parseDateTime = (dateStr: any) => {
    if (!dateStr) return { data: '', hora: '' };
    const parts = dateStr.toString().split(' ');
    const data = parts[0] || '';
    const hora = parts[1] || '';
    return { data, hora };
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setIsLoading(true);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        const processedData = (jsonData as any[]).map((row, idx) => {
          const so = parseDateTime(row['Saída Origem'] || '');
          const co = parseDateTime(row['Chegada Origem'] || '');
          const sd = parseDateTime(row['Saída Destino'] || '');
          const cd = parseDateTime(row['Chegada Destino'] || '');
          
          let mes = '';
          let ano = 0;

          const primarySaida = so.data ? so : (sd.data ? sd : { data: '', hora: '' });

          if (primarySaida.data) {
            const dataParts = primarySaida.data.split('/');
            if (dataParts.length === 3) {
              const mesNum = parseInt(dataParts[1]);
              const meses = ['', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
              mes = meses[mesNum] || '';
              ano = parseInt(dataParts[2]);
            }
          }

          if (!ano) {
            ano = parseInt(row['Ano'] || row['Ano.1'] || 0) || 0;
          }

          // Local normalize status "ConUG" or similar to "Concluído"
          let status = row['Status'] || 'Concluído';
          if (status.toString().toUpperCase().startsWith('CON')) {
            status = 'Concluído';
          }

          return {
            id: row['Id'] || (idx + 1),
            cpf: row['CPF'] || '',
            nome: row['Nome Credor'] || '',
            cargo: row['Cargo'] || '',
            mes,
            ano,
            origem: row['Origem'] || '',
            destino: row['Destino'] || '',
            saidaOrigem: so.data ? `${toISODate(so.data)} ${so.hora}` : '',
            chegadaOrigem: co.data ? `${toISODate(co.data)} ${co.hora}` : '',
            saidaDestino: sd.data ? `${toISODate(sd.data)} ${sd.hora}` : '',
            chegadaDestino: cd.data ? `${toISODate(cd.data)} ${cd.hora}` : '',
            motivo: row['Motivo'] || '',
            status: status,
            totalPago: parseFloat(row['Total Pago'] || 0) || 0
          };
        }).filter(r => r.ano > 0 && r.totalPago >= 0);

        setAllData(processedData);
        
        // Auto set filter to current date if exists
        const hoje = new Date();
        const curMes = MONTH_ORDER[hoje.getMonth()];
        const curAno = hoje.getFullYear();
        if (processedData.some(d => d.mes === curMes && d.ano === curAno)) {
          setMesFilter(curMes);
          setAnoFilter(curAno.toString());
        }

        setIsLoading(false);
      } catch (error) {
        console.error('Error processing file', error);
        alert('Erro ao processar arquivo Excel.');
        setIsLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const filteredData = useMemo(() => {
    return allData.filter(r => {
      const mesMatch = mesFilter === 'all' || r.mes === mesFilter;
      const anoMatch = anoFilter === 'all' || r.ano === parseInt(anoFilter);
      const searchMatch = !searchFilter || 
        r.destino.toLowerCase().includes(searchFilter.toLowerCase()) ||
        r.origem.toLowerCase().includes(searchFilter.toLowerCase()) ||
        r.motivo.toLowerCase().includes(searchFilter.toLowerCase());
      return mesMatch && anoMatch && searchMatch;
    });
  }, [allData, mesFilter, anoFilter, searchFilter]);

  const formatDateSafe = (dateStr: string) => {
    if (!dateStr) return '';
    // Handle both YYYY-MM-DD and DD/MM/YYYY
    if (dateStr.includes('/')) return dateStr;
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    const [year, month, day] = parts;
    return `${day}/${month}/${year}`;
  };

  const formatDateTimeBR = (dateTimeStr: string) => {
    if (!dateTimeStr) return '';
    const parts = dateTimeStr.split(' ');
    const datePart = parts[0];
    const timePart = parts[1];
    
    const formattedDate = formatDateSafe(datePart);
    return timePart ? `${formattedDate} ${timePart}` : formattedDate;
  };

  const toISODate = (val: any): string => {
    if (!val) return '';
    if (typeof val === 'number') {
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    }
    const str = String(val).trim();
    
    // Check if already ISO
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;

    // DD/MM/YYYY or DD/MM/YY
    const ddmmyyyy = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(str);
    if (ddmmyyyy) {
      let [_, d, m, y] = ddmmyyyy;
      if (y.length === 2) y = '20' + y;
      return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }

    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
    return str;
  };

  const getMonthAndYearFromDate = (dateStr: string) => {
    const parts = dateStr.split('-');
    if (parts.length !== 3) return { mes: '', ano: 0 };
    const [year, month] = parts.map(Number);
    return {
      mes: MONTH_ORDER[month - 1],
      ano: year
    };
  };

  const uniqueMeses = useMemo(() => {
    const mesesDiarias = [...new Set(allData.map(r => r.mes).filter(Boolean))] as string[];
    const mesesDespesas = [...new Set(expenses.map(exp => {
      return getMonthAndYearFromDate(exp.date).mes;
    }).filter(Boolean))] as string[];
    const todos = [...new Set([...mesesDiarias, ...mesesDespesas])];
    return todos.sort((a, b) => MONTH_ORDER.indexOf(a) - MONTH_ORDER.indexOf(b));
  }, [allData, expenses]);

  const uniqueAnos = useMemo(() => {
    const anosData = [...new Set(allData.map(r => r.ano).filter(Boolean))] as number[];
    const anosDespesas = [...new Set(expenses.map(exp => getMonthAndYearFromDate(exp.date).ano).filter(Boolean))] as number[];
    const todos = [...new Set([...anosData, ...anosDespesas])];
    return todos.sort((a, b) => a - b);
  }, [allData, expenses]);

  const uniqueCitiesExpenses = useMemo(() => {
    return [...new Set(expenses.map(exp => exp.city).filter(c => c))].sort();
  }, [expenses]);

  const filteredExpenses = useMemo(() => {
    return expenses.filter(exp => {
      const { mes: expMes, ano: expAno } = getMonthAndYearFromDate(exp.date);
      
      const mesMatch = mesFilter === 'all' || expMes === mesFilter;
      const anoMatch = anoFilter === 'all' || expAno === parseInt(anoFilter);
      const cityMatch = cityFilterExpenses === 'all' || exp.city === cityFilterExpenses;
      const typeMatch = !showOnlyAccommodation || exp.type.includes('Hospedagem');
      return mesMatch && anoMatch && cityMatch && typeMatch;
    });
  }, [expenses, mesFilter, anoFilter, cityFilterExpenses, showOnlyAccommodation]);

  const totalPago = useMemo(() => filteredData.reduce((sum, r) => sum + r.totalPago, 0), [filteredData]);
  const totalDespesas = useMemo(() => filteredExpenses.reduce((sum, exp) => sum + exp.value, 0), [filteredExpenses]);
  const totalPernoites = useMemo(() => filteredData.filter(r => r.totalPago > 200).length, [filteredData]);
  const valorLiquido = totalPago - totalDespesas;

  const formatCurrency = (val: number) => {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
  };

  const resetApp = () => {
    setAllData([]);
    setMesFilter('all');
    setAnoFilter('all');
    setSearchFilter('');
    setActiveTab('dados');
  };

  // Chart Data Preparation
  const chartMonthlyData = useMemo(() => {
    const grouped: Record<string, any> = {};
    allData.forEach(r => {
      const key = `${r.ano}-${r.mes}`;
      if (!grouped[key]) grouped[key] = { solicitacoes: 0, pago: 0 };
      grouped[key].solicitacoes++;
      grouped[key].pago += r.totalPago;
    });

    const entries = Object.entries(grouped).map(([key, val]) => {
      const [anoStr, mes] = key.split('-');
      const ano = parseInt(anoStr);
      const despesas = expenses.filter(exp => {
        const { mes: eMes, ano: eAno } = getMonthAndYearFromDate(exp.date);
        return eMes === mes && eAno === ano;
      }).reduce((sum, e) => sum + e.value, 0);

      return {
        key,
        mes,
        ano,
        pago: val.pago,
        solicitacoes: val.solicitacoes,
        despesas,
        liquido: val.pago - despesas
      };
    }).sort((a, b) => {
      if (a.ano !== b.ano) return a.ano - b.ano;
      return MONTH_ORDER.indexOf(a.mes) - MONTH_ORDER.indexOf(b.mes);
    });

    return entries;
  }, [allData, expenses]);

  const chartStatusData = useMemo(() => {
    const grouped: Record<string, number> = {};
    allData.forEach(r => {
      grouped[r.status] = (grouped[r.status] || 0) + 1;
    });
    return grouped;
  }, [allData]);

  if (allData.length === 0 && !isLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-12 rounded-3xl shadow-xl max-w-2xl w-full text-center border border-slate-100"
        >
          <div className="flex justify-center mb-8">
            <div className="p-5 bg-blue-50 rounded-2xl">
              <FileSpreadsheet className="w-16 h-16 text-blue-600" />
            </div>
          </div>
          <h1 className="text-4xl font-bold text-slate-800 mb-4 bg-gradient-to-r from-blue-600 to-indigo-600 bg-clip-text text-transparent">
            Analisador de Diárias
          </h1>
          <p className="text-slate-500 mb-10 text-lg">
            Carregue sua planilha Excel para análise completa de diárias e controle de despesas.
          </p>
          
          <div 
            onClick={() => fileInputRef.current?.click()}
            className="group cursor-pointer border-2 border-dashed border-blue-200 rounded-2xl p-10 bg-blue-50/30 hover:bg-blue-50 hover:border-blue-400 transition-all duration-300"
          >
            <Upload className="w-12 h-12 text-blue-500 mx-auto mb-4 group-hover:scale-110 transition-transform" />
            <p className="text-blue-600 font-semibold text-lg">Clique para selecionar</p>
            <p className="text-slate-400 mt-2">Formatos suportados: .xlsx, .xls</p>
            <input 
              type="file" 
              ref={fileInputRef}
              onChange={handleFileUpload}
              className="hidden" 
              accept=".xlsx,.xls"
            />
          </div>
        </motion.div>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-center">
          <RefreshCw className="w-12 h-12 text-blue-500 animate-spin mx-auto mb-4" />
          <p className="text-slate-600 font-medium">Processando dados...</p>
        </div>
      </div>
    );
  }

  const headerInfo = allData[0] || {};

  return (
    <div className="min-h-screen bg-slate-50 pb-20">
      <header className="sticky top-0 z-40 w-full bg-white/80 backdrop-blur-md border-b border-slate-200">
        <div className="max-w-7xl mx-auto px-4 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center shadow-lg shadow-blue-200">
              <FileSpreadsheet className="text-white w-6 h-6" />
            </div>
            <div className="hidden sm:block">
              <h1 className="text-lg font-black text-slate-800 tracking-tight leading-tight">Remix <span className="text-blue-600 italic">Diárias</span></h1>
              <p className="text-[9px] uppercase tracking-widest font-bold text-slate-400">Analisador de Dados</p>
            </div>
          </div>
          
          <nav className="hidden md:flex bg-slate-100 p-1.5 rounded-2xl border border-slate-200">
            {[
              { id: 'dados', icon: TableIcon, label: 'Diárias' },
              { id: 'despesas', icon: Coins, label: 'Despesas' },
              { id: 'visualizacao', icon: BarChart3, label: 'Gráficos' },
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id as TabType)}
                className={cn(
                  "flex items-center gap-2 px-4 py-2 rounded-xl transition-all duration-300 font-bold text-xs",
                  activeTab === tab.id 
                    ? "bg-white text-blue-600 shadow-sm ring-1 ring-slate-200" 
                    : "text-slate-500 hover:text-slate-800 hover:bg-white/50"
                )}
              >
                <tab.icon className="w-4 h-4" />
                {tab.label}
              </button>
            ))}
          </nav>

          <div className="flex items-center gap-2">
            <button 
              onClick={syncExpenses}
              disabled={isSyncingExpenses}
              className="hidden sm:flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-700 rounded-xl hover:bg-blue-100 transition-colors font-semibold text-xs border border-blue-100 disabled:opacity-50"
            >
              <RefreshCw className={cn("w-3 h-3", isSyncingExpenses && "animate-spin")} />
              Sincronizar
            </button>
            <button 
              onClick={resetApp}
              className="hidden sm:block px-4 py-2 bg-slate-900 text-white rounded-xl hover:bg-slate-800 transition-colors font-semibold text-xs"
            >
              Novo
            </button>
            
            <button 
              onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
              className="md:hidden p-2 text-slate-600 hover:bg-slate-100 rounded-lg"
            >
              <motion.div animate={{ rotate: isMobileMenuOpen ? 90 : 0 }}>
                {isMobileMenuOpen ? <Plus className="w-6 h-6 rotate-45" /> : <Layers className="w-6 h-6" />}
              </motion.div>
            </button>
          </div>
        </div>

        {/* Mobile Navigation */}
        <AnimatePresence>
          {isMobileMenuOpen && (
            <motion.div 
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="md:hidden bg-white border-t border-slate-100 overflow-hidden"
            >
              <div className="px-4 py-6 space-y-2">
                {[
                  { id: 'dados', icon: TableIcon, label: 'Diárias' },
                  { id: 'despesas', icon: Coins, label: 'Despesas' },
                  { id: 'visualizacao', icon: BarChart3, label: 'Gráficos' },
                ].map((tab) => (
                  <button
                    key={tab.id}
                    onClick={() => {
                      setActiveTab(tab.id as TabType);
                      setIsMobileMenuOpen(false);
                    }}
                    className={cn(
                      "flex items-center gap-3 w-full px-4 py-4 rounded-xl font-bold text-base transition-colors",
                      activeTab === tab.id 
                        ? "bg-blue-50 text-blue-600" 
                        : "text-slate-600 hover:bg-slate-50"
                    )}
                  >
                    <tab.icon className="w-5 h-5" />
                    {tab.label}
                  </button>
                ))}
                
                <div className="pt-4 mt-4 border-t border-slate-100 grid grid-cols-2 gap-2">
                  <button 
                    onClick={() => {
                      syncExpenses();
                      setIsMobileMenuOpen(false);
                    }}
                    className="flex items-center justify-center gap-2 py-3 bg-blue-50 text-blue-700 rounded-xl font-bold text-sm"
                  >
                    <RefreshCw className={cn("w-4 h-4", isSyncingExpenses && "animate-spin")} />
                    Sinc
                  </button>
                  <button 
                    onClick={() => {
                      resetApp();
                      setIsMobileMenuOpen(false);
                    }}
                    className="py-3 bg-slate-900 text-white rounded-xl font-bold text-sm"
                  >
                    Novo Arquivo
                  </button>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </header>

      <main className="max-w-7xl mx-auto px-6 mt-8">
        {/* Filters */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
          <div className="flex flex-col gap-1.5">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2 px-1">
              <Calendar className="w-3 h-3" /> Mês
            </label>
            <select 
              value={mesFilter}
              onChange={(e) => setMesFilter(e.target.value)}
              className="w-full bg-white border border-slate-200 rounded-2xl px-4 py-3 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm font-semibold text-slate-700 shadow-sm"
            >
              <option value="all">Filtro: Todos os meses</option>
              {uniqueMeses.map(m => (
                <option key={m} value={m}>{MESES_NUMERO[m] || m}</option>
              ))}
            </select>
          </div>
          <div className="flex flex-col gap-1.5">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2 px-1">
              <Calendar className="w-3 h-3" /> Ano
            </label>
            <select 
              value={anoFilter}
              onChange={(e) => setAnoFilter(e.target.value)}
              className="w-full bg-white border border-slate-200 rounded-2xl px-4 py-3 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm font-semibold text-slate-700 shadow-sm"
            >
              <option value="all">Filtro: Todos os anos</option>
              {uniqueAnos.map(a => (
                <option key={a} value={a}>{a}</option>
              ))}
            </select>
          </div>
          <div className="flex flex-col gap-1.5 sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2 px-1">
              <Search className="w-3 h-3" /> Buscar
            </label>
            <div className="relative">
              <input 
                type="text" 
                placeholder="Pesquisar por destino, motivo..."
                value={searchFilter}
                onChange={(e) => setSearchFilter(e.target.value)}
                className="w-full bg-white border border-slate-200 rounded-2xl pl-11 pr-4 py-3 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm font-semibold text-slate-700 shadow-sm"
              />
              <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4" />
            </div>
          </div>
        </section>

        {/* KPI Cards */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-10">
          {[
            { 
              label: 'Total de Diárias', 
              value: filteredData.length, 
              subValue: `${totalPernoites} pernoite${totalPernoites !== 1 ? 's' : ''}`,
              color: 'from-blue-600 to-blue-700', 
              icon: TableIcon 
            },
            { label: 'Total Recebido', value: formatCurrency(totalPago), color: 'from-emerald-500 to-emerald-600', icon: Coins },
            { label: 'Total Despesas', value: formatCurrency(totalDespesas), color: 'from-rose-500 to-rose-600', icon: Trash2 },
            { label: 'Valor Líquido', value: formatCurrency(valorLiquido), color: 'from-violet-600 to-violet-700', icon: BarChart3 },
          ].map((kpi, i) => (
            <motion.div 
              key={i}
              whileHover={{ y: -4 }}
              className={cn("p-6 rounded-3xl shadow-sm text-white flex flex-col justify-between h-32 bg-gradient-to-br", kpi.color)}
            >
              <div className="flex justify-between items-start">
                <p className="text-base font-bold text-white tracking-wide">{kpi.label}</p>
                <kpi.icon className="w-6 h-6 text-white/40" />
              </div>
              <div className="flex items-baseline justify-between">
                <p className="text-3xl font-black tracking-tight">{kpi.value}</p>
                {('subValue' in kpi) && (
                  <p className="text-xl font-bold tracking-tight text-black">
                    {kpi.subValue}
                  </p>
                )}
              </div>
            </motion.div>
          ))}
        </section>

        {/* Tabs - Now more subtle navigation */}
        <section className="flex gap-1 overflow-x-auto pb-4 scrollbar-hide no-scrollbar -mx-2 px-2 mb-6">
          {[
            { id: 'dados', label: 'Diárias', icon: TableIcon },
            { id: 'despesas', label: 'Despesas', icon: Coins },
            { id: 'analise', label: 'Análise', icon: BarChart3 },
            { id: 'anual', label: 'Anual', icon: Layers },
            { id: 'status', label: 'Status', icon: PieChart },
          ].map((tab) => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id as TabType)}
              className={cn(
                "flex items-center gap-2 px-5 py-3 rounded-2xl font-bold text-xs transition-all whitespace-nowrap shrink-0 border shadow-xs",
                activeTab === tab.id 
                  ? "bg-slate-900 text-white border-slate-900" 
                  : "bg-white text-slate-500 border-slate-200 hover:border-slate-300"
              )}
            >
              <tab.icon className="w-3.5 h-3.5" />
              {tab.label}
            </button>
          ))}
        </section>

        {/* Tab Content */}
        <AnimatePresence mode="wait">
          <motion.div 
            key={activeTab}
            initial={{ opacity: 0, x: 10 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: -10 }}
            transition={{ duration: 0.2 }}
          >
            {activeTab === 'dados' && (
              <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[800px]">
                    <thead>
                      <tr className="bg-slate-50/50 border-b border-slate-100">
                        <th className="px-4 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest w-[15%]">Destino</th>
                        <th className="px-4 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest w-[18%]">Saída/Chegada Or.</th>
                        <th className="px-4 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest w-[18%]">Saída/Chegada Dest.</th>
                        <th className="px-4 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest w-[27%]">Motivo</th>
                        <th className="px-4 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest w-[10%]">Status</th>
                        <th className="px-4 py-4 text-[10px] font-black text-slate-400 uppercase tracking-widest text-right w-[12%]">Valor</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {filteredData.map((r, i) => (
                        <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                          <td className="px-4 py-5 text-sm text-slate-700 font-bold align-top">{r.destino}</td>
                          <td className="px-4 py-5 text-[11px] text-slate-500 leading-tight align-top">
                            <div className="font-bold text-slate-600">S: {formatDateTimeBR(r.saidaOrigem)}</div>
                            <div className="mt-1">C: {formatDateTimeBR(r.chegadaOrigem)}</div>
                          </td>
                          <td className="px-4 py-5 text-[11px] text-slate-500 leading-tight align-top">
                            <div className="font-bold text-slate-600">S: {formatDateTimeBR(r.saidaDestino)}</div>
                            <div className="mt-1">C: {formatDateTimeBR(r.chegadaDestino)}</div>
                          </td>
                          <td className="px-4 py-5 text-xs text-slate-600 leading-relaxed align-top">
                            <div className="line-clamp-4" title={r.motivo}>
                              {r.motivo}
                            </div>
                          </td>
                          <td className="px-4 py-5 align-top">
                            <span className={cn(
                              "px-2 py-1 rounded-full text-[9px] font-black uppercase tracking-tighter block text-center border shadow-xs",
                              r.status === 'Concluído' ? "bg-emerald-50 text-emerald-700 border-emerald-100" : "bg-slate-50 text-slate-500 border-slate-100"
                            )}>
                              {r.status}
                            </span>
                          </td>
                          <td className="px-4 py-5 text-sm font-black text-slate-900 text-right align-top">
                            {formatCurrency(r.totalPago)}
                          </td>
                        </tr>
                      ))}
                      {filteredData.length === 0 && (
                        <tr>
                          <td colSpan={8} className="px-6 py-20 text-center text-slate-400">
                            Nenhum registro encontrado.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {activeTab === 'despesas' && (
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div className="lg:col-span-1 space-y-6">
                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm relative overflow-hidden">
                    <div className="absolute top-0 right-0 p-4 opacity-10">
                      <FileSpreadsheet className="w-24 h-24" />
                    </div>
                    <h3 className="text-lg font-bold text-slate-800 mb-2 flex items-center gap-2">
                       <RefreshCw className={cn("w-5 h-5 text-blue-500", isSyncingExpenses && "animate-spin")} />
                       Status da Planilha
                    </h3>
                    <p className="text-slate-500 text-sm mb-6 leading-relaxed">
                      As despesas agora são sincronizadas automaticamente da Planilha do Google.
                    </p>
                    
                    <div className="space-y-4">
                      <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                        <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Última Sincronização</p>
                        <p className="text-lg font-bold text-slate-800">{lastSyncTime || 'Nunca'}</p>
                        {expenses.length > 0 && (
                          <p className="text-[10px] text-emerald-600 font-bold mt-1 flex items-center gap-1">
                            <RefreshCw className="w-2 h-2" /> Dados carregados com sucesso
                          </p>
                        )}
                      </div>

                      <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                        <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Total de Itens</p>
                        <p className="text-lg font-bold text-slate-800">{expenses.length} registros</p>
                      </div>

                      <a 
                        href="https://docs.google.com/spreadsheets/d/1lyXkSmeiyyODZbng6GtXTSwNR-XY2KWRLKWqppLef1k/edit?gid=980751451#gid=980751451"
                        target="_blank"
                        rel="noopener noreferrer"
                        className="flex items-center justify-center gap-2 w-full py-3 bg-blue-600 text-white rounded-xl font-bold hover:bg-blue-700 transition-colors shadow-lg shadow-blue-100"
                      >
                        <ExternalLink className="w-4 h-4" />
                        Abrir Planilha
                      </a>

                      <button 
                        onClick={syncExpenses}
                        disabled={isSyncingExpenses}
                        className="w-full py-3 text-blue-600 bg-blue-50 hover:bg-blue-100 rounded-xl font-bold transition-colors flex items-center justify-center gap-2 disabled:opacity-50"
                      >
                        <RefreshCw className={cn("w-4 h-4", isSyncingExpenses && "animate-spin")} />
                        Sincronizar Agora
                      </button>
                    </div>
                  </div>

                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                    <h3 className="text-sm font-bold text-slate-400 uppercase mb-4 tracking-widest">Resumo por Tipo</h3>
                    <div className="space-y-3">
                      {Object.entries(
                        filteredExpenses.reduce((acc: any, curr) => {
                          acc[curr.type] = (acc[curr.type] || 0) + curr.value;
                          return acc;
                        }, {})
                      ).map(([type, total]: any) => (
                        <div key={type} className="flex justify-between items-center bg-slate-50 p-3 rounded-xl border border-slate-100">
                          <span className="text-sm text-slate-600 font-medium">{type}</span>
                          <span className="text-sm font-bold text-slate-900">{formatCurrency(total)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="lg:col-span-2 space-y-4">
                  <div className="flex flex-col sm:flex-row justify-between items-center bg-white p-4 rounded-3xl border border-slate-200 shadow-sm gap-4">
                    <div className="flex items-center gap-4">
                      <p className="text-sm font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                         <Search className="w-4 h-4" /> Filtrar
                      </p>
                      <select 
                        value={cityFilterExpenses}
                        onChange={(e) => setCityFilterExpenses(e.target.value)}
                        className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-2 outline-none text-sm focus:ring-2 focus:ring-blue-500 transition-all min-w-[150px]"
                      >
                        <option value="all">Todas as cidades</option>
                        {uniqueCitiesExpenses.map(city => (
                          <option key={city} value={city}>{city}</option>
                        ))}
                      </select>
                    </div>
                    <label className="flex items-center gap-2 cursor-pointer select-none">
                      <input 
                        type="checkbox" 
                        checked={showOnlyAccommodation}
                        onChange={e => setShowOnlyAccommodation(e.target.checked)}
                        className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                      />
                      <span className="text-xs font-bold text-slate-600">Apenas Hotéis/Airbnb</span>
                    </label>
                  </div>

                  {isSyncingExpenses ? (
                    <div className="bg-white p-20 rounded-3xl border border-slate-200 text-center">
                      <RefreshCw className="w-8 h-8 text-blue-500 animate-spin mx-auto mb-4" />
                      <p className="text-slate-500">Sincronizando com Google Sheets...</p>
                    </div>
                  ) : filteredExpenses.length > 0 ? (
                    filteredExpenses.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).map((exp) => (
                      <motion.div 
                        layout
                        key={exp.id} 
                        className="bg-white p-5 rounded-2xl border border-slate-200 flex flex-col sm:flex-row justify-between items-start sm:items-center hover:shadow-md transition-shadow group gap-4"
                      >
                        <div className="flex gap-4 items-start">
                          <div className="w-12 h-12 rounded-xl bg-blue-50 flex items-center justify-center shrink-0">
                            <Coins className="text-blue-600 w-6 h-6" />
                          </div>
                          <div>
                            <div className="flex items-center gap-2 flex-wrap">
                              <p className="text-base font-bold text-slate-900">{exp.city || 'Cidade não informada'}</p>
                              <span className="text-[10px] font-bold bg-slate-100 text-slate-500 px-2 py-0.5 rounded-full uppercase tracking-tight">
                                {exp.type}
                              </span>
                            </div>
                            <p className="text-xs text-slate-400 font-medium my-1">
                              {formatDateSafe(exp.date)} • {exp.description || 'Sem descrição'}
                            </p>
                            <div className="flex items-center gap-3 mt-2">
                              {exp.link && (
                                <a 
                                  href={exp.link} 
                                  target="_blank" 
                                  rel="noopener noreferrer"
                                  className="flex items-center gap-1 text-[11px] font-bold text-blue-600 hover:text-blue-700 hover:underline"
                                >
                                  <ExternalLink className="w-3 h-3" /> Ver Link
                                </a>
                              )}
                              {exp.whatsapp && (
                                <a 
                                  href={`https://wa.me/${exp.whatsapp.replace(/\D/g, '')}`} 
                                  target="_blank" 
                                  rel="noopener noreferrer"
                                  className="flex items-center gap-1 text-[11px] font-bold text-emerald-600 hover:text-emerald-700 hover:underline"
                                >
                                  <MessageCircle className="w-3 h-3" /> WhatsApp
                                </a>
                              )}
                            </div>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 w-full sm:w-auto justify-between sm:justify-end border-t sm:border-t-0 pt-3 sm:pt-0 border-slate-100">
                          <p className="text-lg font-bold text-rose-600">{formatCurrency(exp.value)}</p>
                        </div>
                      </motion.div>
                    ))
                  ) : (
                    <div className="bg-white p-20 rounded-3xl border border-slate-200 text-center text-slate-400">
                      Nenhuma despesa encontrada para os filtros selecionados.
                    </div>
                  )}
                </div>
              </div>
            )}

            {activeTab === 'analise' && (
              <div className="space-y-8">
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                    <h3 className="text-lg font-bold text-slate-800 mb-6">
                      Diárias vs Despesas ({mesFilter === 'all' ? 'Geral' : MESES_NUMERO[mesFilter]} {anoFilter === 'all' ? '' : anoFilter})
                    </h3>
                    {(() => {
                      const targetMes = mesFilter;
                      const targetAno = anoFilter !== 'all' ? parseInt(anoFilter) : null;
                      
                      const filteredDiarios = allData.filter(r => {
                        const mMatch = targetMes === 'all' || r.mes === targetMes;
                        const aMatch = targetAno === null || r.ano === targetAno;
                        return mMatch && aMatch;
                      });

                      const totalDiarias = filteredDiarios.reduce((sum, r) => sum + r.totalPago, 0);
                      
                      const despesasPorTipo: Record<string, number> = {};
                      expenses.filter(exp => {
                        const { mes: expMes, ano: expAno } = getMonthAndYearFromDate(exp.date);
                        const mMatch = targetMes === 'all' || expMes === targetMes;
                        const aMatch = targetAno === null || expAno === targetAno;
                        return mMatch && aMatch;
                      }).forEach(exp => {
                        despesasPorTipo[exp.type] = (despesasPorTipo[exp.type] || 0) + exp.value;
                      });

                      const labels = ['Diárias Recebidas', ...Object.keys(despesasPorTipo)];
                      const dataValues = [totalDiarias, ...Object.values(despesasPorTipo)];
                      const colors = ['#10b981', '#f59e0b', '#3b82f6', '#8b5cf6', '#ef4444', '#06b6d4'];

                      return (
                        <Bar 
                          options={{ 
                            responsive: true, 
                            plugins: { 
                              legend: { display: false },
                              tooltip: {
                                callbacks: {
                                  label: (context) => formatCurrency(context.raw as number)
                                }
                              }
                            } 
                          }}
                          data={{
                            labels: labels,
                            datasets: [
                              { 
                                data: dataValues, 
                                backgroundColor: colors.slice(0, dataValues.length),
                                borderRadius: 8
                              }
                            ]
                          }} 
                        />
                      );
                    })()}
                  </div>
                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                    <h3 className="text-lg font-bold text-slate-800 mb-6">Evolução Mensal (Receita vs Despesa)</h3>
                    <Line 
                      options={{ responsive: true, plugins: { legend: { position: 'bottom' } } }}
                      data={{
                        labels: chartMonthlyData.map(m => `${m.mes} ${m.ano}`),
                        datasets: [
                          { label: 'Recebido', data: chartMonthlyData.map(m => m.pago), borderColor: '#2563eb', backgroundColor: '#2563eb', tension: 0.3 },
                          { label: 'Despesas', data: chartMonthlyData.map(m => m.despesas), borderColor: '#e11d48', backgroundColor: '#e11d48', tension: 0.3 }
                        ]
                      }} 
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                    <h3 className="text-lg font-bold text-slate-800 mb-6">Volume de Solicitações (Últimos 12 Meses)</h3>
                    {(() => {
                      const last12Months = chartMonthlyData.slice(-12);
                      return (
                        <Bar 
                          options={{ responsive: true, plugins: { legend: { display: false } } }}
                          data={{
                            labels: last12Months.map(m => `${m.mes} ${m.ano}`),
                            datasets: [
                              { data: last12Months.map(m => m.solicitacoes), backgroundColor: '#3b82f6', borderRadius: 8 }
                            ]
                          }} 
                        />
                      );
                    })()}
                  </div>
                </div>

                <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead className="bg-slate-50/50 border-b border-slate-100">
                        <tr>
                          <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase">Mês/Ano</th>
                          <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase text-center">Protocolos</th>
                          <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase text-right">Recebido</th>
                          <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase text-right">Despesas</th>
                          <th className="px-6 py-4 text-xs font-bold text-slate-500 uppercase text-right">Líquido</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50">
                        {chartMonthlyData.map((m, i) => (
                          <tr key={i} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-6 py-4 text-sm font-bold text-slate-800 uppercase tracking-wide">{MESES_NUMERO[m.mes]} {m.ano}</td>
                            <td className="px-6 py-4 text-sm text-slate-600 text-center font-medium">{m.solicitacoes}</td>
                            <td className="px-6 py-4 text-sm text-slate-600 text-right">{formatCurrency(m.pago)}</td>
                            <td className="px-6 py-4 text-sm text-rose-600 text-right">{formatCurrency(m.despesas)}</td>
                            <td className={cn("px-6 py-4 text-sm font-bold text-right", m.liquido >= 0 ? "text-emerald-600" : "text-rose-600")}>
                              {formatCurrency(m.liquido)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'anual' && (
              <div className="space-y-8">
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                   <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                    <h3 className="text-lg font-bold text-slate-800 mb-6 text-center">Histórico Anual de Receitas</h3>
                    <div style={{ maxHeight: '350px' }}>
                      <Bar 
                        options={{ 
                          responsive: true, 
                          plugins: { legend: { display: false } },
                          scales: { x: { grid: { display: false } } }
                        }}
                        data={{
                          labels: [...new Set(chartMonthlyData.map(m => m.ano))].sort((a,b) => (a as number) - (b as number)),
                          datasets: [{ 
                             data: [...new Set(chartMonthlyData.map(m => m.ano))].sort((a,b) => (a as number) - (b as number)).map(ano => 
                               chartMonthlyData.filter(m => m.ano === ano).reduce((s, c) => s + c.pago, 0)
                             ),
                             backgroundColor: '#10b981', borderRadius: 12
                          }]
                        }} 
                      />
                    </div>
                  </div>
                  <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                    <h3 className="text-lg font-bold text-slate-800 mb-6 text-center">Histórico Anual de Gastos</h3>
                    <div style={{ maxHeight: '350px' }}>
                      <Bar 
                        options={{ 
                          responsive: true, 
                          plugins: { legend: { display: false } },
                          scales: { x: { grid: { display: false } } }
                        }}
                        data={{
                          labels: [...new Set(chartMonthlyData.map(m => m.ano))].sort((a,b) => (a as number) - (b as number)),
                          datasets: [{ 
                             data: [...new Set(chartMonthlyData.map(m => m.ano))].sort((a,b) => (a as number) - (b as number)).map(ano => 
                               chartMonthlyData.filter(m => m.ano === ano).reduce((s, c) => s + c.despesas, 0)
                             ),
                             backgroundColor: '#f43f5e', borderRadius: 12
                          }]
                        }} 
                      />
                    </div>
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'status' && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-center">
                 <div className="bg-white p-10 rounded-3xl border border-slate-200 shadow-sm flex flex-col items-center">
                   <h3 className="text-lg font-bold text-slate-800 mb-10">Status das Solicitações</h3>
                   <div className="w-full max-w-sm">
                    <Doughnut 
                      data={{
                        labels: Object.keys(chartStatusData),
                        datasets: [{
                          data: Object.values(chartStatusData),
                          backgroundColor: ['#2563eb', '#10b981', '#f59e0b', '#7c3aed', '#f43f5e']
                        }]
                      }}
                      options={{ cutout: '70%', plugins: { legend: { position: 'bottom' } } }}
                    />
                   </div>
                 </div>

                 <div className="space-y-4">
                    {Object.entries(chartStatusData).map(([status, count]) => (
                      <div key={status} className="bg-white p-6 rounded-3xl border border-slate-200 flex justify-between items-center hover:bg-slate-50 transition-colors">
                        <div>
                          <p className="text-sm font-bold text-slate-500 uppercase tracking-widest mb-1">{status}</p>
                          <p className="text-3xl font-black text-slate-800">{count}</p>
                        </div>
                        <div className={cn(
                          "w-3 h-12 rounded-full",
                          status === 'Concluído' ? "bg-emerald-500" : "bg-blue-500"
                        )} />
                      </div>
                    ))}
                 </div>
              </div>
            )}
          </motion.div>
        </AnimatePresence>
      </main>

    </div>
  );
}
