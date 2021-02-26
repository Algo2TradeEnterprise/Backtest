Namespace StrategyHelper
    Public Class RuleEntities
        Enum TargetType
            ATR = 1
            CapitalPercentage
        End Enum

        Enum QuantityType
            Normal = 1
            Increase
        End Enum

        Enum SignalType
            DifferentSignal = 1
            SameSignal
        End Enum

        Enum ReverseType
            Strong = 1
            Any
        End Enum

        Public TypeOfTarget As TargetType
        Public CapitalPercentage As Decimal
        Public TypeOfQuantity As QuantityType
        Public TypeOfSignal As SignalType
        Public TypeOfReverse As ReverseType
        Public TargetLeftPercentage As Decimal
    End Class
End Namespace