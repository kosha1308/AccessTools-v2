Attribute VB_Name = "modATLimitsTests"
Option Explicit

Public Sub AT_TestLimits()
    Dim limits As New ATLimits
    Dim lexed As ATLexResult
    Dim parser As New ATParser
    Dim result As ATParseResult
    Dim compiler As New ATCompiler
    Dim context As New ATCompileContext
    Dim compiled As ATCompileResult
    Dim expression As ATNamedExpression
    Dim context2 As New ATCompileContext
    limits.MaxInputLength = 3
    Set lexed = AT_Tokenize("123", limits)
    AT_Assert lexed.Diagnostics.Count = 0, "input length boundary accepted"
    Set lexed = AT_Tokenize("1234", limits)
    AT_Assert lexed.Diagnostics.Count > 0, "input length boundary plus one rejected"
    limits.MaxInputLength = 100
    limits.MaxTokens = 4
    Set lexed = AT_Tokenize("1+2", limits)
    AT_Assert lexed.Diagnostics.Count = 0, "token limit boundary includes EOF"
    limits.MaxTokens = 3
    Set lexed = AT_Tokenize("1+2", limits)
    AT_Assert lexed.Diagnostics.Count > 0, "token limit exceeded rejected"
    Set limits = New ATLimits
    limits.MaxAstNodes = 3
    Set result = parser.Parse("1+2", "", Nothing, limits)
    AT_Assert result.Diagnostics.Count = 0, "AST node boundary accepted"
    limits.MaxAstNodes = 2
    Set result = parser.Parse("1+2", "", Nothing, limits)
    AT_Assert result.Diagnostics.Count > 0, "AST node limit exceeded rejected"
    Set limits = New ATLimits
    limits.MaxNesting = 8
    Set result = parser.Parse("1+2+3+4+5+6+7+8", "", Nothing, limits)
    AT_Assert result.Diagnostics.Count = 0, "left deep AST boundary accepted"
    Set result = parser.Parse("1+2+3+4+5+6+7+8+9", "", Nothing, limits)
    AT_Assert result.Diagnostics.Count > 0, "left deep AST limited before serialization"
    context.AddField "ROWS", "COLLECTION<ROW<TABLE_A>>"
    Set limits = New ATLimits
    limits.MaxCollectionNesting = 1
    Set compiled = compiler.Compile("ALLROWS(@ROWS AS a, TRUE)", context, "", "BOOL", limits)
    AT_Assert compiled.Diagnostics.Count = 0, "collection nesting boundary accepted"
    Set compiled = compiler.Compile("ALLROWS(@ROWS AS a, ALLROWS(@ROWS AS b, TRUE))", context, "", "BOOL", limits)
    AT_Assert compiled.Diagnostics.Count > 0, "collection nesting exceeded rejected"
    Set expression = New ATNamedExpression
    expression.Initialize "DEFAULT_LOOP", "BOOL", "TRUE"
    expression.AddParameter "P", "BOOL", "$DEFAULT_LOOP"
    context.AddNamed expression
    Set compiled = compiler.Compile("$DEFAULT_LOOP", context)
    AT_Assert compiled.Diagnostics.Count > 0, "recursive default rejected"
    context.AddField "BROKEN_TYPE", "NOT_REGISTERED"
    Set compiled = compiler.Compile("@BROKEN_TYPE = @BROKEN_TYPE", context)
    AT_Assert compiled.Diagnostics.Count > 0, "unknown metadata type rejected"
    Set limits = New ATLimits
    limits.MaxCompileSteps = 1
    Set compiled = compiler.Compile("TRUE", context2, "", "BOOL", limits)
    AT_Assert compiled.Diagnostics.Count = 0, "compile step boundary accepted"
    Set compiled = compiler.Compile("TRUE AND TRUE", context2, "", "BOOL", limits)
    AT_Assert compiled.Diagnostics.Count > 0, "compile step budget exceeded rejected"
    Set expression = New ATNamedExpression
    expression.Initialize "BOTTOM", "BOOL", "TRUE"
    context2.AddNamed expression
    Set expression = New ATNamedExpression
    expression.Initialize "TOP", "BOOL", "$BOTTOM"
    context2.AddNamed expression
    Set limits = New ATLimits
    limits.MaxNamedDepth = 2
    Set compiled = compiler.Compile("$TOP", context2, "", "BOOL", limits)
    AT_Assert compiled.Diagnostics.Count = 0, "named depth boundary accepted"
    limits.MaxNamedDepth = 1
    Set compiled = compiler.Compile("$TOP", context2, "", "BOOL", limits)
    AT_Assert compiled.Diagnostics.Count > 0, "named depth exceeded rejected"
    Set limits = New ATLimits
    limits.MaxNesting = 8
    Set result = parser.Parse("(((((((1)))))))", "", Nothing, limits)
    AT_Assert result.Diagnostics.Count = 0, "syntax nesting boundary accepted"
    Set result = parser.Parse("((((((((1))))))))", "", Nothing, limits)
    AT_Assert result.Diagnostics.Count > 0, "syntax nesting exceeded rejected"
End Sub
