Friend Module Module1

    Sub Main()
        ' Plantilla Ejercicio16ManualVisualBasic2019Jorge
        ' En esta plantilla tendrás lo basico para crear un menu
        ' con un marco que podras ver como se crear mas rapido lento
        ' segundo aumentes o bajes el valor del for en el metodo Temprizador
        'Lo añadido a esta platilla son los metodos por variables
        Ejercicio16ManualVisualBasic2019Jorge()

    End Sub
    Sub Ejercicio16ManualVisualBasic2019Jorge()

        Marco()
        'El menu que haty aqui lo puedes modidicarr a tu gusto 
        'Conserva la numeraracion y sustitulle el resto del texto a tus necesidads
        Console.SetCursorPosition(10, 2)
        Console.Write("Ejercicio16ManualVisualBasic2019Jorge")
        Console.SetCursorPosition(5, 5)
        Console.Write("0.- Salida")
        Console.SetCursorPosition(5, 7)
        Console.Write("1.- Suma por funcion")
        Console.SetCursorPosition(5, 9)
        Console.Write("2. -Numero aleatorio entre un rago de dos numeros dados")
        'Añadrir mas opciones segun vuestras necesidades

        Dim Opcion As Integer
        Dim verdad As Boolean
        ' el sigyiente codigo controla que se introduzca un valor correcto
        Do
            Try
                Console.SetCursorPosition(10, 15)
                Console.Write("elija opcion: ")
                Opcion = Console.ReadLine()
                Eleccion(Opcion)
                verdad = True

            Catch ex As Exception
                Console.SetCursorPosition(10, 17)
                Console.Write("valor introducido no valido")
                verdad = False

            End Try

        Loop Until verdad = True


    End Sub
    Sub Marco()
        Dim Fila As Integer
        Dim Columna As Integer
        For Columna = 0 To 80
            Console.SetCursorPosition(Columna, 0)
            Console.Write("*")
            Temporizador()

        Next
        For Fila = 0 To 21
            Console.SetCursorPosition(80, Fila)
            Console.Write("*")
            Temporizador()

        Next
        For Columna = 80 To 0 Step -1

            Console.SetCursorPosition(Columna, 21)
            Console.Write("*")
            Temporizador()

        Next
        For Fila = 21 To 0 Step -1
            Console.SetCursorPosition(0, Fila)
            Console.Write("*")
            Temporizador()

        Next

    End Sub
    Sub Temporizador()
        ' Este pequello codigo hace la vez de temporizador 
        Dim tempo As Double
        For tempo = 0 To 100000000

        Next
    End Sub
    Sub Eleccion(opcion)
        ' este codigo distribuye el resto de opcines de codigo
        Dim Verdad As Boolean
        Dim numero1 As Double
        Dim numero2 As Double

        Select Case opcion
            Case 0
                Salida()
                Return
            Case 1
                SunaPorFuncion(opcion, Verdad, numero1, numero2)
            Case 2
                Numeroaletorio(opcion, Verdad, numero1, numero2)
            Case Else

                'Una opcion para Controlar  las ociones no aceptadas
                OpcionErronea()
        End Select
    End Sub

    Private Sub Numeroaletorio(opcion As Object, ByRef Verdad As Boolean, ByRef numero1 As Double, ByRef numero2 As Double)
        Borrado()
        Marco()

        Do
            Try
                Console.SetCursorPosition(10, 5)
                Console.Write(" La opcion elegida es: " & opcion)
                Console.SetCursorPosition(10, 7)
                Console.Write("Dime un numero de inicio: ")
                numero1 = Console.ReadLine()
                Console.SetCursorPosition(10, 9)
                Console.Write("Dimeotro numero: ")
                numero2 = Console.ReadLine()
                Console.SetCursorPosition(10, 11)
                Console.Write(" El numero elegido entre: " & numero1 & " y " & numero2 & " es :" & NumAleatorioEntre(numero1, numero2))
                Verdad = True
                Console.ReadKey()

                Borrado()
                Ejercicio16ManualVisualBasic2019Jorge()
            Catch ex As Exception

                Console.SetCursorPosition(10, 20)
                Console.Write(ex.Message)



                Verdad = False


            End Try
        Loop Until Verdad = True

        'Aqui metes nevo codigo
    End Sub

    Private Sub SunaPorFuncion(opcion As Object, ByRef Verdad As Boolean, ByRef numero1 As Double, ByRef numero2 As Double)
        Borrado()
        Marco()

        Do
            Try
                Console.SetCursorPosition(10, 5)
                Console.Write(" La opcion elegida es: " & opcion)
                Console.SetCursorPosition(10, 7)
                Console.Write("Dime un número: ")
                numero1 = Console.ReadLine()
                Console.SetCursorPosition(10, 9)
                Console.Write("Dime otro número: ")
                numero2 = Console.ReadLine()


                Console.SetCursorPosition(10, 11)
                Console.Write("El vañlor de la suma por funcion del valor:" & numero1 & " + " & numero2 & " es:  " & Suma(numero1, numero2))


                Verdad = True
                Borrado()


                Ejercicio16ManualVisualBasic2019Jorge()

            Catch ex As Exception
                Console.SetCursorPosition(10, 20)
                Console.Write(ex.Message)
                Verdad = False


            End Try
        Loop Until Verdad = True


        ' Aqui meter nuevo codigo
    End Sub

    Private Sub OpcionErronea()
        'En  este Metodo controlamos si los numeros introduciodos no son reconocido por el select Case anterior 
        Borrado()
        Marco()


        Console.SetCursorPosition(10, 10)
        Console.Write("OPcion no reconocida")
        Console.ReadKey()
        Borrado()

        Ejercicio16ManualVisualBasic2019Jorge()
    End Sub

    Private Sub Salida()
        Borrado()
        Marco()
        Console.SetCursorPosition(10, 10)
        Console.Write("Gracias pr utilizar la aplicacon")
        Console.ReadKey()
        Borrado()
        End
    End Sub

    Sub Borrado()
        ' este cdigo hace una subifda de pantalla imitando un borrado hasta el efectvo
        Dim Borra As Integer
        For Borra = 0 To 50
            Console.WriteLine()
            Temporizador()

        Next
        Console.Clear()

    End Sub

    Public Function Suma(numero1 As Double, numero2 As Double) As Double

        Suma = numero1 + numero2

    End Function

    Public Function NumAleatorioEntre(ByVal numero1 As Double, ByVal numero2 As Double) As Double
        Randomize(Timer)
        Return CLng((numero1 - numero2) * Rnd() + numero2)


    End Function
End Module
