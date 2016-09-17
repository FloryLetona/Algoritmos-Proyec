Imports System.Text.RegularExpressions



Module Calculadora

    Sub Main()
        Console.Title = "Calculadora de operaciones combinadas"
        Dim operacion As String = ""
        Dim backColor As ConsoleColor = Console.BackgroundColor
        Dim foreColor As ConsoleColor = Console.ForegroundColor

        Console.ForegroundColor = ConsoleColor.White
        Console.BackgroundColor = ConsoleColor.Blue
        Console.BackgroundColor = backColor
        Console.ForegroundColor = foreColor
        Console.WriteLine("Ingrese una operacion combinada o 'end' para salir")

        Do
            operacion = Console.ReadLine()

            If operacion = "end" Then Exit Do

            Try
                ' La salida es en formato Numerico :)
                Console.WriteLine("Resultado: " & Eval(operacion).ToString("N"))
            Catch ex As Exception
                Console.ForegroundColor = ConsoleColor.White
                Console.WriteLine("Error: " & ex.Message())
                Console.ForegroundColor = foreColor
            End Try
        Loop
    End Sub

    ''' <summary>Evalua una ecuacion matematica y la resuelve devolviendo su resultado, como una calculadora cientifica</summary>
    ''' <param name="operacion">Expresion matematica a resolver</param>
    ''' <remarks>Usa Expresiones Regulares :)</remarks>
    ''' <exception cref="FormatException">En caso de que no sea una expresion matematica valida arroja una excepcion</exception>
    Public Function Eval(ByVal operacion As String) As Double
        Dim resp As Double = 0D
        Dim temp As Double = 0D

        ' Reemplazamos los espacios y cambiamos los puntos por comas
        operacion = operacion.Replace(" ", "")
        operacion = operacion.Replace(".", ",")

        Dim RegexObj As New Regex _
            ("(?<Termino>[\+\-]?(?:(?:\d+(?:\,\d*)?|\([\d\,\+\-\/\*]*\))(?:[\*\/](?:\d+(?:\,\d*)?|\([\d\,\+\-\/\*]+\)))*))")

        ' vemos si es una operacion valida
        If RegexObj.IsMatch(operacion) Then
            Dim MatchResults As MatchCollection = RegexObj.Matches(operacion)
            Dim MatchResult As Match = MatchResults(0)
            Dim termino As String


            ' Recorremos los terminos, separados por la Expresion Regular
            For i As Int32 = 0 To MatchResults.Count - 1

                ' Aca Obtenemos "Temp" que es el valor numerico del termino

                termino = MatchResult.Groups("Termino").Value
                If IsNumeric(termino) Then
                    ' Si es numerico, simplemente lo convertimos
                    temp = Double.Parse(termino)
                Else
                    ' En caso de no serlo, significa que tiene signos o parentesis. Aca los tratamos con ResolverTermino

                    ' No podemos pasarle +(5+5)*2 porque lo tomaria como termino y entraria en un bucle infinito
                    ' asi que le extraemos el signo, por defecto le ponemos como positivo y despues le volvemos al
                    ' signo que debe tener
                    Dim signo As Integer = 1

                    If termino.Substring(0, 1) = "-" Then
                        signo = -1
                        termino = termino.Substring(1)
                    ElseIf termino.Substring(0, 1) = "+" Then
                        signo = 1
                        termino = termino.Substring(1)
                    End If

                    temp = ResolverTermino(termino)

                    temp *= signo ' Multiplicamos por -1 para cambiar el signo, por 1 para mantenerlo igual
                End If

                ' Vamos sumando (o restando) y pasamos entre terminos
                resp += temp
                MatchResult = MatchResult.NextMatch()
            Next

            Return resp
        Else
            Throw New FormatException("La operacion no pudo ser reconocida")
        End If
    End Function


    Private Function ResolverTermino(ByVal Termino As String) As Double
        Dim resp As Double = 0D
        Dim temp As Double = 0D

        Dim RegexObj As New Regex( _
            "(?<Termino>[\+\-]? \( [\d\,\+\-\*\/\(\)]+ \)|[\*\/] (?: \( [\d\+\-\*\/\(\)]* \) | \d*\,?\d* )|[\+\-]? \d* (?:\,\d*)?)", _
            RegexOptions.IgnorePatternWhitespace)

        ' Aca Resolvemos potenciacion, multiplicacion y division
        ' Pensaba hacerlo todo en un metodo pero se torno realmente muy dificil, asi que decidi separarlos
        ' por que, "EL ORDEN DE LOS FACTORES NO ALTERA EL PRODUCTO" 
        ' Podia haber expresiones bien complejas, por lo que separamos los grandes terminos de los pequeños

        If RegexObj.IsMatch(Termino) Then

            Dim MatchResults As MatchCollection = RegexObj.Matches(Termino)
            Dim MatchResult As Match = MatchResults(0)
            Dim subTermino As String
            For I As Int32 = 0 To MatchResults.Count - 2
                subTermino = MatchResult.Groups("Termino").Value
                ' Hacemos lo mismo que hacemos en el otro metodo, si es numero, lo devolvemos...
                If IsNumeric(subTermino) Then
                    resp = Double.Parse(subTermino)
                Else
                    ' ... sino, lo identificamos y resolvemos
                    If subTermino.Contains("^") Then
                        Dim regexPotencia As String = "^(?<Base>[+\-]?\d\,?\d*|[+\-]?\([+\-]?\d[\d,+*/\-\^]*\))\^(?<Potencia>[+\-]?\d\,?\d*|[+\-]?\([+\-]?\d[\d,+*/\-\^]*\))$"

                        Dim base As Double = Eval(Regex.Match(subTermino, regexPotencia).Groups("Base").Value)
                        Dim potencia As Double = Eval(Regex.Match(subTermino, regexPotencia).Groups("Potencia").Value)

                        resp = Math.Pow(base, potencia)
                    Else
                        ' Pueden que del otro lado del signo hayan parentesis, asi que lo resolvemos :)
                        Select Case subTermino.Substring(0, 1)
                            Case "*"
                                temp = Eval(subTermino.Substring(1))
                                resp *= temp
                            Case "/"
                                temp = Eval(subTermino.Substring(1))
                                resp /= temp
                            Case Else
                                resp = Eval(Regex.Match(subTermino, "\((?<Operacion>.*)\)").Groups("Operacion").Value)
                        End Select
                    End If
                End If
                MatchResult = MatchResult.NextMatch
            Next
            Return resp
        Else
            Throw New FormatException("Parte de la operacion no pudo ser reconocida")
        End If
    End Function
End Module
