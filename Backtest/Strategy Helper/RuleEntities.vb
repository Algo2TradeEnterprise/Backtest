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

        Public TypeOfTarget As TargetType
        Public CapitalPercentage As Decimal
        Public TypeOfQuantity As QuantityType
    End Class
End Namespace