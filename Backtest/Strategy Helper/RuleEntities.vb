Namespace StrategyHelper
    Public Class RuleEntities
        Enum TargetType
            ATR = 1
            CapitalPercentage
        End Enum

        Public TypeOfTarget As TargetType
        Public CapitalPercentage As Decimal
    End Class
End Namespace