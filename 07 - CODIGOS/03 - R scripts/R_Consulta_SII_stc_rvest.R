

library(rvest)
lego_movie <- html("http://www.imdb.com/title/tt1490017/")

lego_movie %>% 
  html_node("strong span") %>%
  html_text() %>%
  as.numeric()

as.numeric(html_text(html_node(lego_movie,"strong span")))

lego_movie %>%
  html_nodes("#titleCast .itemprop span") %>%
  html_text()

lego_movie %>%
  html_nodes("table") %>%
  .[[3]] %>%
  html_table()


# Para consultar 1 rut
rut <- "5752926"
dv  <- "1"

url_01 <- "https://zeus.sii.cl/cvc_cgi/stc/getstc?RUT="
url_02 <- "&txt_captcha=bUc1Rm5JaHpZYW%20syMDE0MTAxNjE1MzMyMjlBcERZY0hpd2h3MjQyNFZ5b1ZrSktn%20VDhjMDBoSWlsdHhrZ1FqLlFVSk5PR1ZPY1ZGWVl5NUlXUT09em%20RNOVdXWmNVY1E%3D&txt_code=2424&PRG=STC&OPC=NOR"
url_F <- paste(url_01,rut,"&DV=",dv,url_02, sep="")
docXml <- html(url_F)

## # para los Id
## . para los class

html_nodes(docXml,"#barra_sup") %>%
  html_text()

html_node(docXml,"#contenedor div strong") %>%
  html_text()

html_nodes(docXml,"#contenedor")
  html_text()

html_node(docXml,"#contenedor span") %>%
  html_text()

html_nodes(docXml,"#contenedor") %>%
  html_children()


a <-  html_nodes(docXml,"#contenedor") %>% 
        html_children()

html_table(a[25], header = TRUE)

html_node(docXml,"strong") %>%
  html_children()

sections <- html_nodes(docXml, "#contenedor > table ~ table")

html_table(sections[1])

# CSS selectors ----------------------------------------------
ateam <- read_html("http://www.boxofficemojo.com/movies/?id=ateam.htm")
html_nodes(ateam, "center") %>%
  html_text()
html_nodes(ateam, "center font") %>%
  html_text()
html_nodes(ateam, "center font b") %>%
  html_text()

# But html_node is best used in conjunction with %>% from magrittr
# You can chain subsetting:
ateam %>% html_nodes("center") %>% html_nodes("td")
ateam %>% html_nodes("center") %>% html_nodes("font")

td <- ateam %>% html_nodes("center") %>% html_nodes("td")
td
# When applied to a list of nodes, html_nodes() returns all nodes,
# collapsing results into a new nodelist.
td %>% html_nodes("font")
# html_node() returns the first matching node. If there are no matching
# nodes, it returns a "missing" node
if (utils::packageVersion("xml2") > "0.1.2") {
  td %>% html_node("font")
}
