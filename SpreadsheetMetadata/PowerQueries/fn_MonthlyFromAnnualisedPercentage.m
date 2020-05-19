(PercentageAnnuaised as number)=>
let
    MonthlyPercentage = Number.Power((1 + PercentageAnnuaised), (1/12)) -  1
in
    MonthlyPercentage