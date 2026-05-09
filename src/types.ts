export interface DiariaRecord {
  id: string | number;
  cpf: string;
  nome: string;
  cargo: string;
  mes: string;
  ano: number;
  origem: string;
  destino: string;
  saidaOrigem: string;
  chegadaOrigem: string;
  saidaDestino: string;
  chegadaDestino: string;
  motivo: string;
  status: string;
  totalPago: number;
}

export interface Expense {
  id: number;
  date: string;
  value: number;
  type: string;
  description: string;
  city?: string;
  link?: string;
  whatsapp?: string;
}

export type TabType = 'dados' | 'despesas' | 'analise' | 'anual' | 'status';

export const MESES_NUMERO: Record<string, string> = {
  'Jan': 'Janeiro', 'Fev': 'Fevereiro', 'Mar': 'Março', 'Abr': 'Abril',
  'Mai': 'Maio', 'Jun': 'Junho', 'Jul': 'Julho', 'Ago': 'Agosto',
  'Set': 'Setembro', 'Out': 'Outubro', 'Nov': 'Novembro', 'Dez': 'Dezembro'
};

export const MONTH_ORDER = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
