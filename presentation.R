## Workflow
## 1. Read in summarised data such as output from ResQ
## 2. In R Filter the data and specify Triangles and Ultimates data to keep
## 3. Create incurred triangle and ultimates using a triangle data structure
##    either included in ChainLadder or MRMR
## 4. To produce development graphs divide each triangle by corresponding ultimates
## 5. Plot incurred and paid development graphs can be created in several ways
##    Using Plot, and Plot with Lattice = T gives rough and ready default graphs
##    Using ggplot with custom theme gives a polished version

##plot average severity and freqeuency
##plot development graphs + other kpi graphs for specified class

## 6. The Graph object can be output to powerpoint using ReporteRs

library(ReporteRs)
library(ChainLadder)
library(ggplot2)
library(scales)
library(ggthemes)

# Creation of doc, a pptx object (default template)
mydoc = pptx()

##Get some sample data for plotting (RAA dataset)
X <- as.data.frame(RAA)
#X<- as.data.frame(AxQ(x))
#X$Year <- as.factor(X$Origin.Year)
X$Year <- as.factor(X$origin)
p_data <- ggplot(X, aes(dev, value, color=Year))
##The default plot theme
#cc <- scales::seq_gradient_pal("white", "purple", "Lab")(seq(0,1,length.out=10))
p <- p_data +geom_line(linetype="dashed")+geom_point()+ #scale_colour_manual(values=cc)+
    ggtitle("Dataset Development") +
       theme_hc() +
        theme(legend.position="right",
              axis.text=element_text(size=9),
              legend.background = element_rect(fill = "white"),
              legend.text=element_text(color="black"),
              legend.title=element_text(color="black"),
              legend.key=element_rect(size=5)) +
            xlab("Accident Period") + ylab("Incurred Amount")+
                scale_y_continuous(labels = comma) +
                    scale_x_discrete() #+
#                        scale_linetype_manual(breaks=c("1987","1986"), values=c(1,1))
# check my layout names:
slide.layouts(mydoc)

mydoc = addSlide( mydoc, "Two Content" )
# add into mydoc first 10 lines of iris
mydoc = addTitle( mydoc, "A Triangle" )
mydoc = addFlexTable( mydoc, FlexTable(iris[1:10,] ) )

# add text into mydoc (and an empty line just before). Paragraph 
# properties will be those of the shape of the used layout.
mydoc = addParagraph( mydoc, value = c("", "Hello World!") )

mydoc = addSlide( mydoc, "Title and Content" )
# add a plot into mydoc 
mydoc = addPlot( mydoc, function() plot(RAA))
mydoc = addPlot( mydoc, function() print(p))
mydoc = addPlot( mydoc, function() print(plot(AxQ(x),lattice=T))
# write mydoc 
writeDoc( mydoc, "~/pp_simple_example.pptx" )
