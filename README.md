# MACROS.EXCEL
Para realizar macros en excel debes primero 
Abrir Excel despues archivo mas, personalizar obcciones, darle clic en programador, despues ir archvo guardar como guardar macros 
Para iniciar macros. clic en programador y clic en visual basic 

 VBA Es el elnguaje de macros de microsft visual basic que se utiliza  para programar aplicaciones windows y que se incluye en varias aplicaciones  de microsoft 

MACROS Una macro es una aacion o un conjunto de accion que se pueden ejecutar todas las veces que deses .
Si hay tareas de mircrosoft excel que realizas  reiteradamente, puedes grabar una macro para automatizalas 


Diagr de flujo  :  Es una manera visual donde podemos realizar el algoritmo  de forma  de mapa para realizar la solucion del problema 


![Captura](https://user-images.githubusercontent.com/72534486/203465290-125f1986-a814-4b65-a089-1efcb96626c1.PNG)


Variables y tipos de datos  
![image](https://user-images.githubusercontent.com/72534486/203465659-dbfdaae7-52b4-49a9-b516-588ee93dab98.png)


Funciones y subrutinas: 
Procedimientos que realizamos una o varias actividades 
pueden llamar a otros procedimientos
pueden tener o no un parametro
las funciones tienen un valor de retorno
las subrutinas no tienen valor de retorno 

GRABAR MACRO 
Dar clic en programador , dice guardar macros 

 sintaxis
 definir una rutina public/privada  sub nombre parametros as tipo de dato  definir si la funcion o subrutina seraprivada o publica.
 privado se refiere a que solamente se puede utilizar en ese modulo
 publico se refiere q que se puede utilizar en cualquier modulo del documento 
 
Public/Private:
![image](https://user-images.githubusercontent.com/72534486/206873775-49e47f11-b235-43af-8256-c139a9c99381.png)

● Definir si la función o subrutina será privada o pública.
● Privado: se refiere a que solamente se puede utilizar en ese módulo.
● Público: se refiere a que se puede utilizar en cualquier módulo del documento.

Function/Sub:
![image](https://user-images.githubusercontent.com/72534486/206874270-20470758-9ef4-4dd5-a35c-a4c97ce47829.png)

![image](https://user-images.githubusercontent.com/72534486/206874321-68764ab8-69f0-4a59-af0c-4b117d0f1717.png)

● Si se requiere un valor de retorno se utiliza function, si no se necesita valor de retorno, sub.

Nombre:
● Se puede utilizar un nombre que haga referencia al funcionamiento de la función o subrutina.

(parámetros):
● Se incluyen las variables que representan cada uno de los parámetros necesarios.

as Tipo de dato:

● Se escribe el tipo de dato del valor de retorno.
 
Crear una subrutina


 ![image](https://user-images.githubusercontent.com/72534486/203888511-8d1ee44e-f039-4c69-b620-188903bdedc6.png)
 
Crear una subrutina

 ![image](https://user-images.githubusercontent.com/72534486/206874926-f46aabaf-953d-44c2-a653-ffe72ef7b8bd.png)
 Public Sub prueba2()
    
    
    Sheets("Hoja3").Select
    Cells(1, 1) = "hola mundo"
    Range(Cells(2, 2), Cells(3, 3)) = "Hola a todos"
    Range("A3:A5") = "Excel es una locula"
    Range("A3:A5") = 2 * 3
    
End Sub

Message Box

Como incluir un mensaje especifico cuando se cumpla con lactividades requiridas 

![image](https://user-images.githubusercontent.com/72534486/206884954-2a25d17f-4f68-47f5-a185-cde278d5d038.png)

![image](https://user-images.githubusercontent.com/72534486/206885090-586c407a-d5a3-432a-a4d3-389f8f3de26f.png)

excel solucion 
![image](https://user-images.githubusercontent.com/72534486/206885102-9771f271-4650-4e6d-841c-a94b99ebd627.png)

Input Box

Es un mensaje donde podemos crear una ventana emergente 


![image](https://user-images.githubusercontent.com/72534486/206885464-f2d21fef-d0b5-49ad-a99c-8dd2ddb6e380.png)

Public Sub divs()

    Dim num As Integer
    Dim den As Integer
    dim = a
    
    num = InputBox("Numerador", "divicion")
    den = InputBox("denominador", "division")
    
    a = MsgBox(num / den, , "division")


End Sub


Select Case

SE utiliza con las subtutinas en un caso especifico en los casos is es mayor o menor a cero o dar instrucciones es decir si es mayor o menor a cero has esto y si no  has lo otro
es un porgrama para ver mayor o menor con funciones espcificas 


![image](https://user-images.githubusercontent.com/72534486/206885655-69453f55-5200-4aae-b590-52a9b78cb76d.png)

Public Function edad() As Integer
    
    Dim nacimiento As Integer
    Dim a
    Dim b
    Dim año As Integer
    Dim tomar
    año = InputBox("ingrese el año de nacimiento", "Año de nacimiento")
    nacimiento = 2022 - año
    a = MsgBox("tu edad es de " & nacimiento, , "edad")
    Select Case nacimiento
        Case Is >= 19
            tomar = " a tomar"
        Case Else
            tomar = " no puedes tomar"
    End Select
    b = MsgBox("si deseas pero" & tomar, , "salimos")
    
End Function



