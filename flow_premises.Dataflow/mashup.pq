[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
shared data_inicial = #date(2024, 11, 1) meta [IsParameterQuery=true, Type="Date", IsParameterQueryRequired=true];
shared data_final = DateTime.Date(DateTime.LocalNow());
shared base_transacoes = "C:\Users\rob_m\OneDrive - WiseDataTech Consultoria\WiseDataTech\Power BI\GitHub\transacoes.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "calendario_DataDestination", IsNewTarget = true], Settings = [Kind = "Automatic", TypeSettings = [Kind = "Table"]]]}]
shared calendario = let
    dataInicial = data_inicial, 
    dataFinal = data_final, 
    
    datas = List.Dates(
        dataInicial, 
        Duration.Days(dataFinal-dataInicial) + 1, 
        #duration(1, 0, 0, 0)
    ),

    calendario = #table(
        type table[
            Data = date,
            Ano = Int64.Type,
            MesAno = text,
            MesInicio = date
        ],
        List.Transform(
            datas,
            each {
                _,
                Date.Year(_),
                Date.ToText(_, [Format="MMM/yy", Culture="pt-BR"]),
                Date.StartOfMonth(_)
            }
        )
    )

in
    calendario;
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "transacoes_DataDestination", IsNewTarget = true], Settings = [Kind = "Automatic", TypeSettings = [Kind = "Table"]]]}]
shared transacoes = let
  Origem = Excel.Workbook(Web.Contents("https://onetechbicombr-my.sharepoint.com/personal/robson_wisedatatech_com_br/Documents/WiseDataTech/Power%20BI/GitHub/transacoes.xlsx"), null, true),
  #"Navegação 1" = Origem{[Item = "transacoes", Kind = "Sheet"]}[Data],
  #"Cabeçalhos promovidos" = Table.PromoteHeaders(#"Navegação 1", [PromoteAllScalars = true]),
  #"Tipo de coluna alterado" = Table.TransformColumnTypes(#"Cabeçalhos promovidos", {{"ID", Int64.Type}, {"Data", type date}, {"Total", type number}, {"Moeda", type text}}),
  #"Linhas filtradas" = Table.SelectRows(#"Tipo de coluna alterado", each [Data] >= data_inicial)
in
  #"Linhas filtradas";
shared transacoes_DataDestination = let
  Pattern = Lakehouse.Contents([HierarchicalNavigation = null, CreateNavigationProperties = false, EnableFolding = false]),
  Navigation_1 = Pattern{[workspaceId = "71e4fce6-e93b-476b-a294-530a37e86ff6"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "8f08d04c-e443-4e34-b72d-98e47bbe49f8"]}[Data],
  TableNavigation = Navigation_2{[Id = "transacoes", ItemKind = "Table"]}?[Data]?
in
  TableNavigation;
shared calendario_DataDestination = let
  Pattern = Lakehouse.Contents([HierarchicalNavigation = null, CreateNavigationProperties = false, EnableFolding = false]),
  Navigation_1 = Pattern{[workspaceId = "71e4fce6-e93b-476b-a294-530a37e86ff6"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "8f08d04c-e443-4e34-b72d-98e47bbe49f8"]}[Data],
  TableNavigation = Navigation_2{[Id = "calendario", ItemKind = "Table"]}?[Data]?
in
  TableNavigation;
