import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';

// Props do componente PedidosFerias
export interface IPedidosFeriasProps {
  listName: string;
  context: WebPartContext;
  sp: SPFI;
  graph: GraphFI;
}
