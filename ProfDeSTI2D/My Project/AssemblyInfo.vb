Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' Les informations générales relatives à un assembly dépendent de 
' l'ensemble d'attributs suivant. Changez les valeurs de ces attributs pour modifier les informations
' associées à un assembly.

' Vérifiez les valeurs des attributs de l'assembly

<Assembly: AssemblyTitle("Prof de STI2D")>
<Assembly: AssemblyDescription("Ruban d'aide à la création de documents STI2D.")>
<Assembly: AssemblyCompany("Electropol")>
<Assembly: AssemblyProduct("Prof de STI2D")>
<Assembly: AssemblyCopyright("CopyLeft  2021")>
<Assembly: AssemblyTrademark("")> 

' L'affectation de la valeur false à ComVisible rend les types invisibles dans cet assembly 
' aux composants COM.  Si vous devez accéder à un type dans cet assembly à partir de 
' COM, affectez la valeur true à l'attribut ComVisible sur ce type.
<Assembly: ComVisible(False)>

'Le GUID suivant est pour l'ID de la typelib si ce projet est exposé à COM
<Assembly: Guid("1fe0be7f-cb35-4dd4-9f69-8440a9627817")>

' Les informations de version pour un assembly se composent des quatre valeurs suivantes :
'
'      Version principale
'      Version secondaire 
'      Numéro de build
'      Révision
'
' Vous pouvez spécifier toutes les valeurs ou utiliser par défaut les numéros de build et de révision 
' en utilisant '*', comme indiqué ci-dessous :
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("2.0.0.0")>
<Assembly: AssemblyFileVersion("2.0.0.0")>

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
