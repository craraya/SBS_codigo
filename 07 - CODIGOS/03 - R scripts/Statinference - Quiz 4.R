
# Question 1


ncp<-(1100)*sqrt(9)/30

delta<-qt(p=0.95, df=8, ncp=(1100)*sqrt(9)/30, lower.tail = TRUE, log.p = FALSE)

1100-delta
1100+delta
