import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import { GraphFI, graphfi, SPFx as GraphSPFx } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/graph/users';

import * as strings from 'PedidosFeriasWebPartStrings';
import PedidosFerias from './components/PedidosFerias';
import { IPedidosFeriasProps } from './components/IPedidosFeriasProps';

export interface IPedidosFeriasWebPartProps {
  listName: string;
}

// WebPart principal - gere pedidos de ferias dos colaboradores
export default class PedidosFeriasWebPart extends BaseClientSideWebPart<IPedidosFeriasWebPartProps> {

  private _sp: SPFI;
  private _graph: GraphFI;

  // Inicializa o PnPjs quando a webpart carrega
  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
    this._graph = graphfi().using(GraphSPFx(this.context));
  }

  // Renderiza o componente React
  public render(): void {
    const element: React.ReactElement<IPedidosFeriasProps> = React.createElement(
      PedidosFerias,
      {
        listName: this.properties.listName || 'Pedidos de Férias',
        context: this.context,
        sp: this._sp,
        graph: this._graph
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Configuracoes que aparecem no painel lateral da webpart
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  value: this.properties.listName || 'Pedidos de Férias'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
