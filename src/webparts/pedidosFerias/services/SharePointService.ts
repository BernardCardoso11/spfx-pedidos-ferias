import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IPedidoFerias,
  IPedidoFeriasListItem,
  EstadoPedido,
  OrdenacaoTipo
} from '../models';

// Serviço para operações com a lista SharePoint
export class SharePointService {
  private _sp: SPFI;
  private _listName: string;

  constructor(sp: SPFI, listName: string) {
    this._sp = sp;
    this._listName = listName;
  }

  // Obtém os pedidos da lista
  public async getPedidos(
    ordenacao: OrdenacaoTipo = 'desc',
    colaboradorId?: number
  ): Promise<IPedidoFerias[]> {
    try {
      // Constrói a query base
      let query = this._sp.web.lists
        .getByTitle(this._listName)
        .items
        .select(
          'Id',
          'Title',
          'ColaboradorId',
          'Colaborador/Id',
          'Colaborador/Title',
          'Colaborador/EMail',
          'DataInicio',
          'DataFim',
          'Estado',
          'Observacoes',
          'Created',
          'Modified'
        )
        .expand('Colaborador')
        .orderBy('DataInicio', ordenacao === 'asc');

      // Adiciona filtro por colaborador se especificado
      if (colaboradorId) {
        query = query.filter(`ColaboradorId eq ${colaboradorId}`);
      }

      // Executa a query
      const items: IPedidoFeriasListItem[] = await query();

      // Mapeia os resultados para o modelo IPedidoFerias
      return items.map(item => this.mapToIPedidoFerias(item));
    } catch (error) {
      console.error('Erro ao obter pedidos de férias:', error);
      throw new Error(`Falha ao carregar pedidos de férias: ${error.message}`);
    }
  }

  // Atualiza o estado de um pedido
  public async atualizarEstado(
    pedidoId: number,
    novoEstado: EstadoPedido
  ): Promise<void> {
    try {
      await this._sp.web.lists
        .getByTitle(this._listName)
        .items
        .getById(pedidoId)
        .update({
          Estado: novoEstado
        });
    } catch (error) {
      console.error(`Erro ao atualizar estado do pedido ${pedidoId}:`, error);
      throw new Error(`Falha ao atualizar estado: ${error.message}`);
    }
  }

  // Aprova um pedido
  public async aprovarPedido(pedidoId: number): Promise<void> {
    await this.atualizarEstado(pedidoId, 'Aprovado');
  }

  // Rejeita um pedido
  public async rejeitarPedido(pedidoId: number): Promise<void> {
    await this.atualizarEstado(pedidoId, 'Rejeitado');
  }

  // Mapeia o item da lista para o nosso modelo
  private mapToIPedidoFerias(item: IPedidoFeriasListItem): IPedidoFerias {
    return {
      Id: item.Id,
      Title: item.Title,
      ColaboradorId: item.ColaboradorId,
      ColaboradorNome: item.Colaborador?.Title || 'Desconhecido',
      ColaboradorEmail: item.Colaborador?.EMail || '',
      DataInicio: item.DataInicio,
      DataFim: item.DataFim,
      Estado: item.Estado,
      Observacoes: item.Observacoes,
      Created: item.Created,
      Modified: item.Modified
    };
  }
}

export default SharePointService;
