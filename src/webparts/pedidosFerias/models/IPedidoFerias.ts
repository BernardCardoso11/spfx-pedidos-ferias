// Representa um pedido de férias
export interface IPedidoFerias {
  Id: number;
  Title?: string;
  ColaboradorId: number;
  ColaboradorNome?: string;
  ColaboradorEmail?: string;
  DataInicio: string;
  DataFim: string;
  Estado: EstadoPedido;
  Observacoes?: string;
  Created?: string;
  Modified?: string;
}

// Estados possíveis de um pedido
export type EstadoPedido = 'Pendente' | 'Aprovado' | 'Rejeitado';

// Colaborador expandido do SharePoint
export interface IColaborador {
  Id: number;
  Title: string;
  EMail: string;
}

// Item da lista SharePoint (antes de mapear)
export interface IPedidoFeriasListItem {
  Id: number;
  Title?: string;
  ColaboradorId: number;
  Colaborador?: IColaborador;
  DataInicio: string;
  DataFim: string;
  Estado: EstadoPedido;
  Observacoes?: string;
  Created?: string;
  Modified?: string;
}

// Tipo de ordenação
export type OrdenacaoTipo = 'asc' | 'desc';
