
#Question 2
pnorm(70, mean = 80, sd = 10, lower.tail = TRUE, log.p = FALSE)

#Question 3
qnorm(0.95, mean = 1100, sd = 75, lower.tail = TRUE, log.p = FALSE)

#Question 4
#rnorm(100, mean = 1100, sd = 75)
qnorm(0.95, mean = 1100, sd = 75/sqrt(100), lower.tail = TRUE, log.p = FALSE)

#Question 5
0.5^4*0.5 +.5^5

#Question 6
pnorm(16, mean = 15, sd = 1, lower.tail = TRUE, log.p = FALSE) -
pnorm(14, mean = 15, sd = 1, lower.tail = TRUE, log.p = FALSE)

#Question 8
ppois(10, 5*3, lower.tail = TRUE, log.p = FALSE)
  
