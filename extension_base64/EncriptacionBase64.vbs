'******************************************************************************
'* File:     Encriptacion Base64
'* Purpose:  Generar los Opaques en Base64 para agregarlos a los Secrets K8S
'* Title:    Encriptacion Base64
'* Author:   Clodoaldo Sanchez Perez
'* Creación: Nov 03, 2019
'* Version:  1.0
'* Comment:  
'******************************************************************************
'Crear la Extensión: Encriptacion
'' + Profile
'' - Class
''   - Menus
''     - Menu_Encriptar
''
''      <Menu>
''         <Command Name="Encriptar_Base64" Caption="Encriptar Base64" />
''      </Menu>
''       
''     - Methods
''       - Encriptar_Base64
'******************************************************************************

Dim Hasclass 'Class existence in active selection
Dim obj
Dim diagram 'the current diagram
Dim strresult

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
            output "Base64 de Atributos Class: "& obj.Code
            output "----------------------------------------------"
            strresult = getOpaques(obj)
            
            Output strresult
            output "----------------------------------------------"
         End If
      Next
      If not Hasclass Then
         output "No class selected"
      End If
   End If

End Sub

' Esta funcion devuelve el campo Name en Mayuscula y
' Utilizo esta funcion para sombrear el valor del comment.
' y queda listo para agregarlo a los secrets de kubernetes
Function getOpaques(clase)
Dim resultado
Dim ATT
For Each ATT in clase.Attributes
   'output ATT.code
   'output ATT.comment
   resultado = resultado + space(4) + Ucase(ATT.Name) + " : " + Base64Encode(ATT.comment) + vbNewLine 
Next
getOpaques = resultado
end Function

' Las funciones de encriptacion base64 fueron obtenidas de:
' https://www.motobit.com/tips/detpg_Base64Encode/
Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, I
  
  'For each group of 3 bytes
  For I = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _
      &H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function