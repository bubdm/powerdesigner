# POWERDESIGNER OOM - ENCRIPTACION BASE64

Codigo para Generar Opaques desde PowerDesigner para Secrets K8s a partir
de los atributos de una clase:

ClassName: Debe contener el Nombre del Secret
Comments : Debe contener el Valor del Secret a Encriptar

## Instalación

    ******************************************************************************
    Crear la Extensión: Encriptacion
     + Profile
     - Class
       - Menus
         - Menu_Encriptar
    
          <Menu>
             <Command Name="Encriptar_Base64" Caption="Encriptar Base64" />
          </Menu>
           
         - Methods
           - Encriptar_Base64
    ******************************************************************************

[CSP]: 2019
