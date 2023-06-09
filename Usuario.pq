let
    Source = Excel.Workbook(File.Contents(Ruta&"\UsuariosEmsa.xlsx"), null, true),
    Usuario_Table = Source{[Item="Usuario",Kind="Table"]}[Data],
    Replaced_Fecha_Min = Table.ReplaceValue(Usuario_Table,null, each Limites_De_Fechas{0}[FechaMinima],Replacer.ReplaceValue,{"Fecha Inicio"}),
    Replaced_Fecha_Max = Table.ReplaceValue(Replaced_Fecha_Min,null, each Limites_De_Fechas{0}[FechaMaxima],Replacer.ReplaceValue,{"Fecha Fin"}),
    Changed_Type = Table.TransformColumnTypes(Replaced_Fecha_Max,{{"Acceso", type text}, {"Nombre", type text}, {"Usuario", type text}, {"Clave", Int64.Type}, {"FormulaClave", Int64.Type}, {"Fecha Inicio", type date}, {"Fecha Fin", type date}, {"Nivel Acceso", type text}}),
    Merged_Util_Usuario = Table.NestedJoin(Changed_Type, {"Acceso"}, UtilUsuario, {"Acceso"}, "UtilUsuario", JoinKind.LeftOuter),
    Expanded_Util_Usuario = Table.ExpandTableColumn(Merged_Util_Usuario, "UtilUsuario", {"Nivel_Acceso"}, {"Nivel_Acceso"})
in
    Expanded_Util_Usuario
