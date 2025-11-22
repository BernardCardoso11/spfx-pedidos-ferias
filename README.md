# Pedidos de Ferias - WebPart SPFx

## O que faz

- Lista os pedidos de ferias de uma lista SharePoint
- Filtra por colaborador (People Picker)
- Ordena por data de inicio
- Permite aprovar ou rejeitar pedidos (atualiza o campo Estado)
- Mostra a foto do colaborador via Graph API

## Como usar

### 1. Instalar dependencias

```bash
npm install
```

### 2. Configurar a lista SharePoint

Lista chamada "Pedidos de Ferias" com estas colunas:

- **Colaborador** (Person)
- **DataInicio** (Date)
- **DataFim** (Date)
- **Estado** (Choice: Pendente, Aprovado, Rejeitado)
- **Observacoes** (Multiple lines)

### 3. Correr localmente

```bash
gulp serve
```

### 4. Build para producao

```bash
gulp bundle --ship
gulp package-solution --ship
```

## Autor

Bernardo Cardoso
