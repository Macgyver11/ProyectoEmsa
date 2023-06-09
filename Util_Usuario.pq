// UtilUsuario
let
    Source = Excel.Workbook(File.Contents(Ruta&"\UsuariosEmsa.xlsx"), null, true),
    Usuario_Table = Source{[Item="Usuario",Kind="Table"]}[Data],
    Changed_Type = Table.TransformColumnTypes(Usuario_Table,{{"Acceso", type text}, {"Nombre", type text}, {"Usuario", type text}, {"Clave", Int64.Type}, {"FormulaClave", Int64.Type}, {"Fecha Inicio", type date}, {"Fecha Fin", type date}, {"Nivel Acceso", type text}}),
    Selected_Columns = Table.SelectColumns(Changed_Type,{"Acceso", "Nivel Acceso"}),
    Grouped_Rows = Table.Group(Selected_Columns, {"Nivel Acceso"}, {{"All", each _, type table [Acceso=nullable text, Nivel Acceso=nullable text]}}),
    Added_Custom = Table.AddColumn(Grouped_Rows, "Custom", each Table.AddColumn(Table.AddIndexColumn([All], "Index", 1, 1, Int64.Type), "Nivel_Acceso", each Text.Combine({[Nivel Acceso], Text.From([Index], "es-ES")}, "_"), type text), type table[Acceso = text, Nivel Acceso = text, Index = number, Nivel_Acceso = text]),
    Expanded_Custom = Table.ExpandTableColumn(Added_Custom, "Custom", {"Acceso", "Nivel_Acceso"}, {"Acceso", "Nivel_Acceso"}),
    Removed_Other_Columns = Table.SelectColumns(Expanded_Custom,{"Nivel Acceso", "Acceso", "Nivel_Acceso"})
in
    Removed_Other_Columns
