
require(Rfacebook)
library(devtools)
library(ggplot2)

# https://developers.facebook.com/tools/explorer/

fb_oauth <- fbOAuth(app_id="10209565917043931")

token<-'EAACEdEose0cBAB8SUKF6EViBpGk9KbgjhYSx6AFdM3TIcYYKnMLg6b92isT0fZBjReHZCAduwWzF3RzL97DkTEO3oh9zJh1IBAQuct2jakofZCuPrZBUS0AzZCJHksd8tjmvGgZBH1GVMzIZBJndslgKntegkMGtmb5DGUChu5XkOPEjbraxOeZBukSDQMgc984ZD'

me<-getUsers("carlos.araya.8", token=token, private_info=TRUE)
names(me)
me$name

getLikes("me",token)[1,]

my_friends <- getFriends(token)

head(my_friends$id, n = 1) # get lowest user ID

gg<-getUsers("carlos.araya.8", token)
names(gg)
gg$picture

## desde paginas
page2 <- getPage("humansofnewyork", token, n = 1000)

page <- getPage("misifuz88", token, n = 500)
page <- getFriends("misifuz88", token, n = 500)

#page[which.max(page$likes_count), ]

format.facebook.date <- function(datestring) {
  date <- as.POSIXct(datestring, format = "%Y-%m-%dT%H:%M:%S+0000", tz = "GMT")
}
## aggregate metric counts over month
aggregate.metric <- function(metric) {
  m <- aggregate(page[[paste0(metric, "_count")]], list(month = page$month), 
                 mean)
  m$month <- as.Date(paste0(m$month, "-15"))
  m$metric <- metric
  return(m)
}
# create data frame with average metric counts per month
page$datetime <- format.facebook.date(page$created_time)
page$month <- format(page$datetime, "%Y-%m")
df.list <- lapply(c("likes", "comments", "shares"), aggregate.metric)
df <- do.call(rbind, df.list)
# visualize evolution in metric
library(ggplot2)
library(scales)
ggplot(df, aes(x = month, y = x, group = metric)) + geom_line(aes(color = metric)) + 
  scale_x_date(date_breaks = "years", labels = date_format("%Y")) + scale_y_log10("Average count per post", 
                                                                                  breaks = c(10, 100, 1000, 10000, 50000)) + theme_bw() + theme(axis.title.x = element_blank())



