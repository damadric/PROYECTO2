Public Module ReglasDeInactivacion 
    Public Function SeDebeInactivar(laFechaDeUltimaTransaccion As Date, laFechaActual As Date) As Boolean 
        Const losDiasPermitidosDeInactividad = 30 
        Dim seInactiva As Boolean 
        Dim losDiasDeInactividad As Integer 
        losDiasDeInactividad = ObtengaLosDias(laFechaDeUltimaTransaccion, laFechaActual) 
    
        If losDiasDeInactividad >= losDiasPermitidosDeInactividad Then 
            seInactiva = True 
        Else 
            seInactiva = False 
        End If 
        Return seInactiva 
    End Function 
    
    Private Function ObtengaLosDias(laFechaDeUltimaTransaccion As Date, laFechaActual As Date) As Double
        Return(laFechaActual - laFechaDeUltimaTransaccion).TotalDays
    End Function

End Module