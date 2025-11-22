import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  DefaultButton,
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  Icon,
  Persona,
  PersonaSize,
  IPersonaProps
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

import styles from './PedidosFerias.module.scss';
import { IPedidosFeriasProps } from './IPedidosFeriasProps';
import { IPedidoFerias, OrdenacaoTipo, EstadoPedido } from '../models';
import { SharePointService, GraphService } from '../services';

// Componente para gerir pedidos de férias
const PedidosFerias: React.FC<IPedidosFeriasProps> = (props) => {
  const { listName, context, sp, graph } = props;

  // Estados do componente
  const [pedidos, setPedidos] = useState<IPedidoFerias[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [erro, setErro] = useState<string | null>(null);
  const [ordenacao, setOrdenacao] = useState<OrdenacaoTipo>('desc');
  const [colaboradorFiltro, setColaboradorFiltro] = useState<number | undefined>(undefined);
  const [fotos, setFotos] = useState<Map<string, string>>(new Map());
  const [atualizando, setAtualizando] = useState<number | null>(null);

  // Instâncias dos serviços
  const spService = new SharePointService(sp, listName);
  const graphService = new GraphService(graph);

  // Carrega os pedidos da lista
  const carregarPedidos = useCallback(async () => {
    try {
      setLoading(true);
      setErro(null);

      // Obtém os pedidos da lista
      const resultado = await spService.getPedidos(ordenacao, colaboradorFiltro);
      setPedidos(resultado);

      // Carrega as fotos dos colaboradores (opcional)
      const emails = resultado
        .map(p => p.ColaboradorEmail)
        .filter(e => e) as string[];

      if (emails.length > 0) {
        const fotosCarregadas = await graphService.getFotosUtilizadores(emails);
        setFotos(fotosCarregadas);
      }
    } catch (error) {
      console.error('Erro ao carregar pedidos:', error);
      setErro(error.message || 'Erro ao carregar pedidos de férias');
    } finally {
      setLoading(false);
    }
  }, [ordenacao, colaboradorFiltro, listName]);

  // Carrega os pedidos quando o componente monta ou quando os filtros mudam
  useEffect(() => {
    carregarPedidos();
  }, [carregarPedidos]);

  // Aprova um pedido
  const handleAprovar = async (pedidoId: number): Promise<void> => {
    try {
      setAtualizando(pedidoId);
      await spService.aprovarPedido(pedidoId);

      // Atualiza o estado local
      setPedidos(prev =>
        prev.map(p =>
          p.Id === pedidoId ? { ...p, Estado: 'Aprovado' as EstadoPedido } : p
        )
      );
    } catch (error) {
      setErro(`Erro ao aprovar pedido: ${error.message}`);
    } finally {
      setAtualizando(null);
    }
  };

  // Rejeita um pedido
  const handleRejeitar = async (pedidoId: number): Promise<void> => {
    try {
      setAtualizando(pedidoId);
      await spService.rejeitarPedido(pedidoId);

      // Atualiza o estado local
      setPedidos(prev =>
        prev.map(p =>
          p.Id === pedidoId ? { ...p, Estado: 'Rejeitado' as EstadoPedido } : p
        )
      );
    } catch (error) {
      setErro(`Erro ao rejeitar pedido: ${error.message}`);
    } finally {
      setAtualizando(null);
    }
  };

  // Quando o utilizador seleciona alguém no People Picker
  const handlePeoplePickerChange = (items: IPersonaProps[]): void => {
    if (items.length > 0) {
      // Extrai o ID do utilizador selecionado
      const userId = (items[0] as any).id;
      setColaboradorFiltro(userId ? parseInt(userId, 10) : undefined);
    } else {
      setColaboradorFiltro(undefined);
    }
  };

  // Quando muda a ordenação
  const handleOrdenacaoChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      setOrdenacao(option.key as OrdenacaoTipo);
    }
  };

  // Formata a data para pt-PT
  const formatarData = (data: string): string => {
    if (!data) return '-';
    return new Date(data).toLocaleDateString('pt-PT', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    });
  };

  // Calcula quantos dias de férias
  const calcularDias = (dataInicio: string, dataFim: string): number => {
    if (!dataInicio || !dataFim) return 0;
    const inicio = new Date(dataInicio);
    const fim = new Date(dataFim);
    const diffTime = Math.abs(fim.getTime() - inicio.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1; // +1 para incluir o dia inicial
    return diffDays;
  };

  // Devolve a classe CSS do estado
  const getEstadoCor = (estado: EstadoPedido): string => {
    switch (estado) {
      case 'Aprovado':
        return styles.estadoAprovado;
      case 'Rejeitado':
        return styles.estadoRejeitado;
      default:
        return styles.estadoPendente;
    }
  };

  // Colunas da tabela
  const colunas: IColumn[] = [
    {
      key: 'colaborador',
      name: 'Colaborador',
      fieldName: 'ColaboradorNome',
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
      onRender: (item: IPedidoFerias) => (
        <Persona
          imageUrl={fotos.get(item.ColaboradorEmail || '') || undefined}
          text={item.ColaboradorNome}
          size={PersonaSize.size32}
          imageAlt={`Foto de ${item.ColaboradorNome}`}
        />
      )
    },
    {
      key: 'dataInicio',
      name: 'Início',
      fieldName: 'DataInicio',
      minWidth: 80,
      maxWidth: 90,
      isResizable: true,
      onRender: (item: IPedidoFerias) => (
        <Text>{formatarData(item.DataInicio)}</Text>
      )
    },
    {
      key: 'dataFim',
      name: 'Fim',
      fieldName: 'DataFim',
      minWidth: 80,
      maxWidth: 90,
      isResizable: true,
      onRender: (item: IPedidoFerias) => (
        <Text>{formatarData(item.DataFim)}</Text>
      )
    },
    {
      key: 'numeroDias',
      name: 'Dias',
      minWidth: 40,
      maxWidth: 50,
      isResizable: true,
      onRender: (item: IPedidoFerias) => (
        <Text>{calcularDias(item.DataInicio, item.DataFim)}</Text>
      )
    },
    {
      key: 'estado',
      name: 'Estado',
      fieldName: 'Estado',
      minWidth: 90,
      maxWidth: 110,
      isResizable: true,
      onRender: (item: IPedidoFerias) => (
        <Text className={getEstadoCor(item.Estado)}>
          <Icon
            iconName={
              item.Estado === 'Aprovado'
                ? 'CheckMark'
                : item.Estado === 'Rejeitado'
                ? 'Cancel'
                : 'Clock'
            }
          />{' '}
          {item.Estado}
        </Text>
      )
    },
    {
      key: 'acoes',
      name: 'Ações',
      minWidth: 180,
      maxWidth: 200,
      onRender: (item: IPedidoFerias) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          {item.Estado === 'Pendente' && (
            <>
              <PrimaryButton
                text="Aprovar"
                iconProps={{ iconName: 'Accept' }}
                onClick={() => handleAprovar(item.Id)}
                disabled={atualizando === item.Id}
                className={styles.botaoAprovar}
                styles={{ root: { minWidth: 80, padding: '0 8px' } }}
              />
              <DefaultButton
                text="Rejeitar"
                iconProps={{ iconName: 'Cancel' }}
                onClick={() => handleRejeitar(item.Id)}
                disabled={atualizando === item.Id}
                className={styles.botaoRejeitar}
                styles={{ root: { minWidth: 80, padding: '0 8px' } }}
              />
            </>
          )}
          {item.Estado !== 'Pendente' && (
            <Text variant="small" className={styles.estadoFinal}>
              {item.Estado === 'Aprovado' ? '✓ Aprovado' : '✗ Rejeitado'}
            </Text>
          )}
        </Stack>
      )
    }
  ];

  // Opções de ordenação
  const ordenacaoOpcoes: IDropdownOption[] = [
    { key: 'desc', text: 'Decrescente' },
    { key: 'asc', text: 'Crescente' }
  ];

  return (
    <div className={styles.pedidosFerias}>
      {/* Cabeçalho */}
      <div className={styles.cabecalho}>
        <Text variant="xLarge" className={styles.titulo}>
          <Icon iconName="Calendar" /> Pedidos de Férias
        </Text>
      </div>

      {/* Filtros */}
      <div className={styles.filtros}>
        <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
          {/* People Picker para filtrar por colaborador */}
          <Stack.Item grow>
            <Text variant="small" className={styles.labelFiltro}>
              Filtrar por colaborador:
            </Text>
            <PeoplePicker
              context={context as any}
              personSelectionLimit={1}
              showtooltip={true}
              required={false}
              onChange={handlePeoplePickerChange}
              principalTypes={[PrincipalType.User]}
              resolveDelay={200}
              placeholder="Selecione um colaborador..."
              webAbsoluteUrl={context.pageContext.web.absoluteUrl}
              ensureUser={true}
              searchTextLimit={2}
            />
          </Stack.Item>

          {/* Dropdown para ordenação */}
          <Stack.Item>
            <Text variant="small" className={styles.labelFiltro}>
              Ordenar por data de início:
            </Text>
            <Dropdown
              selectedKey={ordenacao}
              options={ordenacaoOpcoes}
              onChange={handleOrdenacaoChange}
              styles={{ dropdown: { width: 200 } }}
            />
          </Stack.Item>

          {/* Botão para recarregar */}
          <Stack.Item align="end">
            <DefaultButton
              text="Atualizar"
              iconProps={{ iconName: 'Refresh' }}
              onClick={carregarPedidos}
              disabled={loading}
            />
          </Stack.Item>
        </Stack>
      </div>

      {/* Mensagem de erro */}
      {erro && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          onDismiss={() => setErro(null)}
          dismissButtonAriaLabel="Fechar"
          className={styles.mensagemErro}
        >
          {erro}
        </MessageBar>
      )}

      {/* Loading */}
      {loading && (
        <div className={styles.loading}>
          <Spinner size={SpinnerSize.large} label="A carregar pedidos..." />
        </div>
      )}

      {/* Lista de pedidos */}
      {!loading && pedidos.length > 0 && (
        <DetailsList
          items={pedidos}
          columns={colunas}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
          className={styles.lista}
        />
      )}

      {/* Mensagem quando não há pedidos */}
      {!loading && pedidos.length === 0 && !erro && (
        <MessageBar messageBarType={MessageBarType.info}>
          Não foram encontrados pedidos de férias.
          {colaboradorFiltro && ' Tente remover o filtro de colaborador.'}
        </MessageBar>
      )}

      {/* Rodapé com estatísticas */}
      {!loading && pedidos.length > 0 && (
        <div className={styles.rodape}>
          <Text variant="small">
            Total: {pedidos.length} pedido(s) |
            Pendentes: {pedidos.filter(p => p.Estado === 'Pendente').length} |
            Aprovados: {pedidos.filter(p => p.Estado === 'Aprovado').length} |
            Rejeitados: {pedidos.filter(p => p.Estado === 'Rejeitado').length}
          </Text>
        </div>
      )}
    </div>
  );
};

export default PedidosFerias;
