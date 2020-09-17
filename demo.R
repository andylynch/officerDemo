library(ggplot2)
library(officer)
library(ggplot2)

doc<- read_docx("officerDemoInput.docx")
cat("Founds these bookmarks: ", officer::docx_bookmarks(doc))

# Prepare some important data

opinions <- data.frame(label= paste("Pie charts are" , 
 c( "confusing","hard too read"," my favourite", "just wrong")), value= sample.int(100, 4))

# And a nice chart
plot<- ggplot(opinions, aes(x="", y=value, fill=label)) +
  geom_bar(stat="identity", width=1) +
  coord_polar("y", start=0)

cursor_begin(doc) # Move to start of the document

cursor_reach(doc,"table placeholder")
officer::body_remove(doc) # remove placeholder 
officer::body_add_table(doc, opinions)

cursor_reach(doc,"chart placeholder")
officer::body_remove(doc) # remove placeholder 
body_add_gg(doc,plot)

if(!dir.exists("output")){ dir.create("output")}
print(doc,file.path("output","result.docx"))
