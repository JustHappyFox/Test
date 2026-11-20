// Ищет условные выражения a ? b : c, где типы b и c неявно приводятся друг к другу (byte и int и т.п.):
// старый компилятор C# на Windows 7 их не принимает. Аргументы: файл .cs, затем справочные сборки.
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
var src = File.ReadAllText(args[0]);
var tree = CSharpSyntaxTree.ParseText(src, new CSharpParseOptions(LanguageVersion.CSharp2));
var refs = args.Skip(1).Select(p => (MetadataReference)MetadataReference.CreateFromFile(p)).ToList();
var comp = CSharpCompilation.Create("x", new[] { tree }, refs, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
var model = comp.GetSemanticModel(tree);
int n = 0;
foreach (var c in tree.GetRoot().DescendantNodes().OfType<ConditionalExpressionSyntax>()) {
    var a = model.GetTypeInfo(c.WhenTrue).Type; var b = model.GetTypeInfo(c.WhenFalse).Type;
    if (a == null || b == null || SymbolEqualityComparer.Default.Equals(a, b)) continue;
    if (a.TypeKind == TypeKind.Error || b.TypeKind == TypeKind.Error) continue;
    // ни один не приводится неявно к другому как ТИП - или оба (по константе): считаем приведения по выражениям
    bool ab = model.ClassifyConversion(c.WhenTrue, b).IsImplicit, ba = model.ClassifyConversion(c.WhenFalse, a).IsImplicit;
    if (ab && ba) { n++; var l = c.GetLocation().GetLineSpan().StartLinePosition.Line + 1; Console.WriteLine($"{args[0]}:{l}: {a} и {b}: {c}"); }
}
Console.WriteLine($"?: с двусторонним приведением: {n}");
