setwd("c:\\USERS\\WONHEE")

x<-read.csv("mapL.csv", header = T)
x <- x[,c(1,2)]
names(x) <- c("Date", "Price")

x$Date = as.Date(x$Date)

sample <- x[x$Date >= as.Date('2016-03-26') & 
         x$Date <= as.Date('2016-10-14'),]	
		
sample$Price = log(sample$Price)
sample$Dummy = 0


sample$Dummy = ifelse(sample$Date >= as.Date('2016-06-26') &
                      sample$Date <= as.Date('2016-07-14') , 1, 0)

DummyReg <- lm(Price ~ Date + Dummy, data = sample, method = "qr",
               x = TRUE, y = TRUE)

summary(DummyReg)
plot(sample$Date, sample$Price, type = 'l')
