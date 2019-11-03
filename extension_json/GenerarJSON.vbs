'******************************************************************************
'* File:     Generar JSON
'* Purpose:  Crear el objeto JSON de una Clase en Power Designer
'* Title:    Notacion JSON
'* Author:   Clodoaldo Sanchez Perez
'* Creación: Nov 02, 2019
'* Version:  1.0
'* Comment:  
'******************************************************************************
'Crear la Extensión: Notacion JSON
'' + Profile
'' - Class
''   - Menus
''     - Menu_Generar_JSON
''
''      <Menu>
''         <Command Name="Generar_JSON" Caption="Generar JSON" />
''      </Menu>
''       
''     - Methods
''       - Generar_JSON
'******************************************************************************

Dim Hasclass 'Class existence in active selection
Dim obj
Dim diagram 'the current diagram
Dim strJson
dim asocProcesadas

Set asocProcesadas = CreateObject("System.Collections.ArrayList")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Metodo Principal a ser llamado desde el Menu Contextual de la Clase.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub %Method%(obj)
   ' Implement your method on <obj> here
   Output "Inicio"
   
   Dim nb 'Number of object in active selection
   nb = ActiveSelection.Count
   If nb = 0 Then 
       output " No seleccionó clases"
   Else
      Hasclass = false
      For Each obj in ActiveSelection
         If obj.IsKindOf(PdOOM.cls_Class)  Then
            Hasclass = True
            Set diagram = ActiveDiagram
            output "----------------------------------------------"
            output "JSON de Class: "& obj.Code
            output "----------------------------------------------"
            'Transform obj, diagram
            asocProcesadas.clear
            strJson = generaJson( obj, 1, "", false)
            Output strJson
            output "----------------------------------------------"
         End If
      Next
      If not Hasclass Then
         output "No class selected"
      End If
   End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Funciòn de Generación
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function generaJson(clss, nivel, nodo, esArray)
Dim resultado
Dim ATT
dim ASOC
dim skipLocal
dim esArrayLocal
Dim procesada
procesada = false
esArrayLocal = false
skipLocal = 0
if nodo ="" then
   resultado = "{"
else
   if esArray then
      resultado = space(nivel * 2) + """"+camelCase(nodo)+""":"+vbNewLine + space(nivel * 2) +"[{" +vbNewLine
   else
      resultado = space(nivel * 2) + """"+camelCase(nodo)+""":"+vbNewLine + space(nivel * 2) +"{" +vbNewLine
   end if
end if
'output "nivel === "+ CStr(nivel)
'output "generaJson: " + clss.code
'output "nodo: " + nodo
   
   For Each ATT in clss.Attributes
      'clss.Attributes.Remove(ATT)
      'myinterface.Attributes.Add(ATT)
      'output ATT.code
      'output ATT.dataType
      resultado = resultado + space(nivel * 2 + 1) + formatCampo(ATT.code, ATT.dataType) + vbNewLine 
      procesada = true
   Next
   if procesada then
      asocProcesadas.add clss
   end if
   For Each ASOC in clss.Associations
      'output "------asoc--->" + asoc.code
      if fueProcesada(asoc.classB) then
         'no hacemos nada
      else
'             output "------asoc--->" + asoc.code
'             output "------asoc--->" + asoc.classB.Name
'             output "------asoc--->" + asoc.RoleBName
'             output "------asoc--->" + asoc.RoleBContainer
         'output space(nivel * 2) + formatCampo(ASOC.RoleBName, ASOC.RoleBContainer, asoc.classB)
         if asoc.RoleBContainer = "java.util.List" then
            esArrayLocal = true
         end if
         resultado = resultado + generaJson( asoc.classB, nivel+1, asoc.RoleBName, esArrayLocal) + vbNewLine 
      end if
   Next

if esArray then
   resultado = resultado + space(nivel * 2) + "}],"
else
   resultado = resultado + space(nivel * 2) + "},"
end if

generaJson = resultado

end function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function fueProcesada(clase)
'output "buscando->" +clase
Dim resultado 
Dim cls
resultado = false
For Each cls in asocProcesadas
   if clase = cls then
      resultado = true
'      output "..encontrado->" +clase
   end if
next
fueProcesada = resultado
end function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function formatCampo(campo, tipo)
Dim strTipo
'output strTipo
select case tipo
   case "java.util.Date", "Date"
      strTipo = "ISODate("+"""2019-10-31T22:27:07.226-0500"""+")"
      'strTipo = """2019-10-31T22:27:07.226-0500"""
   case "java.lang.String", "String"
      strTipo = """"""
   case "BigDecimal", "float", "java.lang.Double", "Double"
      strTipo = "0.00"
   case "java.lang.Number", "Number", "double"
      strTipo = "0.00"
   case "java.lang.Long", "Long", "java.lang.Integer", "Integer"
      strTipo = "0"
   case "int"
      strTipo = "0"
   case "Boolean", "java.lang.Boolean", "boolean"
      strTipo = "true"
   case "java.lang.Object", "Object"
      strTipo = "0"
   case "java.util.List", "List"
      strTipo = "[]"
   case else
      Output "revisar estension " + tipo
      strTipo = "valor"
end select

formatCampo = """"+camelCase(campo)+""""+":"+strTipo+","
end function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function camelCase(cadena)
Dim ancho
ancho = len(cadena)
camelCase = lcase(mid(cadena, 1, 1)) + right(cadena, ancho - 1)
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
