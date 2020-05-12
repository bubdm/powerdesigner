'******************************************************************************
'* File:     Generador de Diccionario de Datos
'* Purpose:  Generar un Diccionario de Datos en Excel de un Archivo PDM
'* Title:    Diccionario de Datos
'* Author:   Clodoaldo Sanchez Perez
'* Creación: May 08, 2014
'* Version:  1.0
'* Comment:  Abrir un Modelo Fìsico de Datos y Ejecutar el Script con la opcion
'*           "Run Command" del Menú
'******************************************************************************

Option Explicit

'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------

Dim nb
dim gi_fila_ini, gi_fila_fin
Dim lb_Ancho_ajustado
' Get the current active model
Dim Model

lb_Ancho_ajustado = False
Set Model = ActiveModel
If (Model Is Nothing) Or (Not Model.IsKindOf(PdPDM.cls_Model)) Then
   MsgBox "The current model is not an PDM model."
Else
   ' Get the Tables collection
   Dim ModelTables
   Set ModelTables = Model.Tables
   Output ". El Modelo '" + Model.Name + "' Contiene " + CStr(ModelTables.Count) + " Tablas."
   Output ""
   Output ". Inicio del Proceso. "
   
   ' Open & Create Excel Document
	Dim x1  '
	Set x1 = CreateObject("Excel.Application") 
	x1.Visible = True 
	x1.Workbooks.Add
	nb = 2

   ShowProperties Model
   Output ". Fin del Proceso. "
End If

'-----------------------------------------------------------------------------
' Show properties of Package
'-----------------------------------------------------------------------------
Sub ShowProperties(package)
   ' Show tables of the current model/package
   Dim noTable
   noTable = 1
   ' For each Table
   Dim tbl
   For Each tbl In package.Tables
      ShowTable tbl, noTable
      noTable = noTable + 1
   Next
      
End Sub

'-----------------------------------------------------------------------------
' Show Table properties
'-----------------------------------------------------------------------------
Sub ShowTable(tbl, noTable)
   If IsObject(tbl) Then
      Dim bShortcutClosed
      bShortcutClosed = false
      If tbl.IsShortcut Then
         If tbl.Status = "Opened" Then
            ' Show properties of the target table
            Set tbl = tbl.TargetObject
         Else
            ' The target model is not opened (closed or not found)
            bShortcutClosed = true
         End If
      End If 

      If Not bShortcutClosed Then
      	'x1.Range("A"+Cstr(nb)).Value = tbl.Name
         Output ".. Procesando Tabla "+tbl.Code + ". "
         gi_fila_ini = nb 'Guardo la Posicion Inicial
         
         x1.Range("A"+Cstr(nb)).Value = "Código de Tabla"
         x1.Range("A"+Cstr(nb)).Interior.ColorIndex = 15 'Gris
         x1.Range("B"+Cstr(nb)).Value = tbl.Code
         x1.Range("B"+Cstr(nb)+":G"++Cstr(nb)).Merge
         nb = nb + 1
         x1.Range("A"+Cstr(nb)).Value = "Nombre de Tabla"
         x1.Range("A"+Cstr(nb)).Interior.ColorIndex = 15 'Gris
         x1.Range("B"+Cstr(nb)).Value = tbl.Name
         x1.Range("B"+Cstr(nb)+":G"++Cstr(nb)).Merge
         nb = nb + 1
         x1.Range("A"+Cstr(nb)).Value = "Descripción de Tabla"
         x1.Range("A"+Cstr(nb)).Interior.ColorIndex = 15
         'Modificacion CSP 19/05/2014
         If len(tbl.Comment) = 0 or isnull(tbl.Comment) then
            x1.Range("B"+Cstr(nb)).Value = tbl.Name
         else
            x1.Range("B"+Cstr(nb)).Value = tbl.Comment
         end if
         'Fin Modificacion
         x1.Range("B"+Cstr(nb)+":G"++Cstr(nb)).Merge
         nb = nb + 1
         x1.Range("A"+Cstr(nb)).Value = "Definición de Campos"
         x1.Range("A"+Cstr(nb)+":G"++Cstr(nb)).Merge
         x1.Range("A"+Cstr(nb)).Interior.ColorIndex = 15
         x1.Range("A"+Cstr(nb)).HorizontalAlignment = -4108 'Centrado
         nb = nb + 1
      	x1.Range("A"+Cstr(nb)).Value = "Código                 "
      	x1.Range("B"+Cstr(nb)).Value = "Nombre                                  "	
       	x1.Range("C"+Cstr(nb)).Value = "Tipo de Dato         "
       	x1.Range("D"+Cstr(nb)).Value = "Valores                                 "    
       	x1.Range("E"+Cstr(nb)).Value = "Nulo (S/N)   "    
       	x1.Range("F"+Cstr(nb)).Value = "PK    "
       	x1.Range("G"+Cstr(nb)).Value = "FK    "
         
         x1.Range("A"+Cstr(nb)+":G"+Cstr(nb)).Interior.ColorIndex = 15 'Gris

         nb = nb + 1

      	' Show columns
      	ShowColumns tbl

         If Not lb_Ancho_ajustado Then
       	  x1.Columns("A:G").EntireColumn.AutoFit 'Ajustar el Ancho de columnas
            lb_Ancho_ajustado = True
         End if

         x1.Range("A"+Cstr(nb)).Value = "Indices"
         x1.Range("A"+Cstr(nb)).HorizontalAlignment = -4108 'Centrado
         With x1.Range("A"+Cstr(nb)+":G"++Cstr(nb))
            .Merge
            .Interior.ColorIndex = 15 'Gris
            With .Borders
               .LineStyle = 1 'xlContinuous
               .Color = vbBlack
               .Weight = 2 'xlThin
            End With
         end With
         nb = nb + 1
      	x1.Range("A"+Cstr(nb)).Value = "Nombre               "
      	x1.Range("B"+Cstr(nb)).Value = "Columnas             "	
         With x1.Range("B"+Cstr(nb)+":G"++Cstr(nb))
            .Merge
         end With
         x1.Range("A"+Cstr(nb)+":G"++Cstr(nb)).Interior.ColorIndex = 15 'Gris
         With x1.Range("A"+Cstr(nb)+":G"++Cstr(nb)).Borders
            .LineStyle = 1 'xlContinuous
            .Color = vbBlack
            .Weight = 2 'xlThin
         end With
         nb = nb + 1

         ShowIndexes tbl
         
         gi_fila_fin = nb -2
         'Colocar Lineas
         With x1.Range("A"+Cstr(gi_fila_ini )+":G"++Cstr(gi_fila_fin)).Borders
            .LineStyle = 1 'xlContinuous
            .Color = vbBlack
            .Weight = 2 'xlThin
      end With
         gi_fila_ini = 0
         gi_fila_fin = 0
      Else
         Output "The target table of the shortcut " + tbl.Code + " is not accessible."
         Output ""
      End If
   End If
End Sub

'-----------------------------------------------------------------------------
' Show Table columns
'-----------------------------------------------------------------------------
Sub ShowColumns(tbl)
   If IsObject(tbl) Then
      Dim col
      For Each col In tbl.Columns
         If Not col.IsShortcut Then
            x1.Range("A"+Cstr(nb)).Value = col.Code
            x1.Range("B"+Cstr(nb)).Value = col.Name
            'Modificacion CSP 19/05/2014
            x1.Range("B"+Cstr(nb)).WrapText = True
            'Fin Modificacion CSP 19/05/2014
            'x1.Range("B"+Cstr(nb)).WrapText = True
            x1.Range("C"+Cstr(nb)).Value = col.DataType
            'Modificacion CSP 19/05/2014
            If len(col.Comment) = 0 or isnull(col.Comment) then
               x1.Range("D"+Cstr(nb)).Value = col.Name
               if col.Code = col.Name then
                  x1.Range("D"+Cstr(nb)).Interior.ColorIndex = 3 'Rojo
               end if
            else
               x1.Range("D"+Cstr(nb)).Value = col.Comment
            end if
            'Fin Modificacion                     
            x1.Range("D"+Cstr(nb)).WrapText = True
            If col.Mandatory then
               x1.Range("E"+Cstr(nb)).Value = "S"
            Else
               x1.Range("E"+Cstr(nb)).Value = "N"
            End if
       		If col.Primary Then
       			x1.Range("F"+Cstr(nb)).Value = "X"
       		Else
       			x1.Range("F"+Cstr(nb)).Value = ""
       		End If
       		If col.ForeignKey Then
       			x1.Range("G"+Cstr(nb)).Value = "X"
       		Else
       			x1.Range("G"+Cstr(nb)).Value = ""
       		End If
           	nb = nb + 1
         End If
      Next
      'nb = nb + 1 'Salto al final de procesar las Columnas
   End If
End Sub

'Mostrar Los Indices
Sub ShowIndexes(tbl)
Dim li_fila_ini_idx, li_fila_fin_idx
Dim ls_col_Valor, li_pos_punto
   If IsObject(tbl) Then
      Dim Indice
      for each Indice in tbl.indexes
         If Not Indice.IsShortcut Then
            x1.Range("A"+Cstr(nb)).Value = Indice.Code
            Dim ColIndex
            li_fila_ini_idx = nb
            For each ColIndex in Indice.IndexColumns
               li_pos_punto = InStr(ColIndex.Column , ".")
               ls_col_Valor = mid(ColIndex.Column, li_pos_punto +1, len(ColIndex.Column) - li_pos_punto  -1)
			     x1.Range("B"+Cstr(nb)).Value = ls_col_Valor
               x1.Range("B"+Cstr(nb)+":G"+Cstr(nb)).merge
               nb = nb + 1 
            next
            ColIndex = null
            li_fila_fin_idx = nb -1
            If li_fila_fin_idx > li_fila_ini_idx Then
               x1.Range("A"+Cstr(li_fila_ini_idx)+":A"+Cstr(li_fila_fin_idx)).merge
            end if
         End if
      next
      nb = nb + 1 'Salto al final de procesar las Columnas
   End If
End Sub