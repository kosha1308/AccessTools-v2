Attribute VB_Name = "modATAppendixTests"
Option Explicit

Public Sub AT_TestAppendix()
    Dim context As New ATCompileContext, compiler As New ATCompiler
    Dim expression As ATNamedExpression, result As ATCompileResult
    Dim item As Variant
    context.AddField "TECH_DEST_PORT", "PORT_SET"
    context.AddField "SOURCE_PORT", "PORT_SET"
    context.AddField "TECH_DEST_IP", "IP_EXPRESSION"
    context.AddField "ALLOWED_IP_RANGE", "IP_EXPRESSION"
    context.AddField "TECH_PROTOCOL", "TEXT"
    context.AddField "REQUEST_TYPE", "TEXT"
    context.AddField "QUANTITY", "NUMBER"
    context.AddField "PRICE", "NUMBER"
    context.AddField "INITIATOR_EMAIL", "EMAIL"
    context.AddField "REQUESTED_DEPARTMENT", "TEXT"
    context.AddField "TECHNICAL_ROWS", "COLLECTION<ROW<TECHNICAL>>"
    context.AddField "ADDRESS_ROWS", "COLLECTION<ROW<ADDRESS>>"
    context.AddField "TECH_PROTOCOL", "TEXT", "TECHNICAL"
    context.AddField "TECH_DEST_PORT", "PORT_SET", "TECHNICAL"
    context.AddField "TECH_DEST_IP", "IP_EXPRESSION", "TECHNICAL"
    context.AddField "ADDRESS_IP", "IP_EXPRESSION", "ADDRESS"
    context.AddSource "ALLOWED_NETWORKS", "IP_EXPRESSION"
    context.AddSource "INTERNAL_NETWORKS", "IP_EXPRESSION"
    Set expression = New ATNamedExpression
    expression.Initialize "PORT_ALLOWED", "BOOL", "@TECH_DEST_PORT COVER $PORTS"
    expression.AddParameter "PORTS", "PORT_SET"
    context.AddNamed expression
    Set expression = New ATNamedExpression
    expression.Initialize "IS_STANDARD_WEB", "BOOL", "@TECH_DEST_PORT COVER: 443,8443"
    context.AddNamed expression
    Set expression = New ATNamedExpression
    expression.Initialize "IS_CONTRACTOR", "BOOL", "TRUE"
    context.AddNamed expression
    For Each item In Array("COVER: 443,8443 AND NOT ANY: 9000", "(COVER: 443,8443 AND NOT ANY: 9000) OR EXACT: 9443", "@TECH_DEST_PORT = @SOURCE_PORT", "@TECH_DEST_IP CONTAINED_BY @ALLOWED_IP_RANGE", "(@TECH_PROTOCOL = ""TCP"" AND @TECH_DEST_PORT ANY: 443,8443) OR (@TECH_PROTOCOL = ""UDP"" AND @TECH_DEST_PORT EXACT: 53)", "COUNT(@TECHNICAL_ROWS) > 10", "@QUANTITY * @PRICE > 100000", "@TECH_DEST_IP IN SOURCE(""ALLOWED_NETWORKS"")", "$IS_STANDARD_WEB", "$PORT_ALLOWED([443,8443])")
        Set result = compiler.Compile(CStr(item), context, "TECH_DEST_PORT")
        AT_Assert Not result.Root Is Nothing, "appendix N scalar/source/named sample"
    Next item
    For Each item In Array("ALLROWS(@TECHNICAL_ROWS AS t, @t.TECH_PROTOCOL = ""TCP"" AND @t.TECH_DEST_PORT COVER: 443,8443)", "ANYROW(@TECHNICAL_ROWS AS t, @t.TECH_PROTOCOL = ""TCP"" AND @t.TECH_DEST_PORT ANY: 443)", "ALLROWS(@TECHNICAL_ROWS AS t, ANYROW(@ADDRESS_ROWS AS a, @a.ADDRESS_IP INTERSECTS @t.TECH_DEST_IP))", "COUNT(FILTER(@TECHNICAL_ROWS AS t, @t.TECH_PROTOCOL = ""TCP"" AND @t.TECH_DEST_PORT ANY: 22,3389)) > 3", "@REQUEST_TYPE = ""Create"" AND $IS_CONTRACTOR AND @TECH_DEST_IP CONTAINED_BY SOURCE(""INTERNAL_NETWORKS"") AND NOT @TECH_DEST_PORT EXACT: 443 AND ANYROW(@TECHNICAL_ROWS AS t, @t.TECH_PROTOCOL = ""TCP"")")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert Not result.Root Is Nothing, "appendix N collection/correlated sample"
    Next item
    Set result = compiler.Compile("DISTINCT_COUNT(@TECH_DEST_IP) > 5", context)
    AT_Assert result.Root Is Nothing And result.Diagnostics.Count > 0, "appendix aggregate requires registered collection binding"
    Set result = compiler.Compile("LOOKUP(""EMPLOYEE_DIRECTORY"", @INITIATOR_EMAIL, ""Department"") = @REQUESTED_DEPARTMENT", context)
    AT_Assert result.Root Is Nothing And result.Diagnostics.Count > 0, "appendix unregistered lookup adapter fails closed"
    Set result = compiler.Compile("ALLROWS: 10.1.1.1,10.1.1.2", context)
    AT_Assert result.Root Is Nothing And result.Diagnostics.Count > 0, "legacy ALLROWS requires explicit migration importer"
End Sub
