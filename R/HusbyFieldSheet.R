

HusbyFieldSheet <- function(path){
  library(readxl)
  library(dplyr)
  library(stringr)
  
  
  SheetNames<- excel_sheets(path)
  lengthLoop <- length(SheetNames)-11
  
  for (i in 1:lengthLoop){
    #Getting the first row of excel sheet "H_M_C2_018"
    PlotDateInventorLocation <- read_excel(
      path,
      sheet = SheetNames[i],
      skip = 0,
      n_max = 1,
      col_names = F)
    
    #Replace NA with 0
    PlotDateInventorLocation[is.na(PlotDateInventorLocation)] <- 0
    
    #Get every other column
    col_names<-PlotDateInventorLocation[seq(1,length(PlotDateInventorLocation),2)]
    
    #converting to list
    PlotDateInventorLocation<-PlotDateInventorLocation[1,]
    
    #checking for missing values for "sted"
    if (PlotDateInventorLocation[8]==0){
      PlotDateInventorLocation[8] <- "missing"
    }
    
    #Getting the Plot, Date, Inventor and Location columns
    df<-col_names[1:4]
    colnames(df) <- col_names[1:4]
    
    #Making the date column as date type (maybe set the value to NA first, if it doesn't work)
    df$Dato <- as.Date(df$Dato, format = "%Y.%m.%d")
    
    
    #Inserting values of Plot, Date, Inventor and Location
    for(j in 1:4){
      if (PlotDateInventorLocation[4]==0){
        df$Dato<-NA
        df[1,1]<-PlotDateInventorLocation[1*2]
        df[1,2]<-NA
        df[1,3]<-PlotDateInventorLocation[3*2]
        df[1,4]<-PlotDateInventorLocation[4*2]
      } else {
        df[1,j]<-PlotDateInventorLocation[j*2]
      }
      
    }
    
    
    
    
    
    
    #Importing rows with "blomster"
    AllData <- read_excel(
      path,
      sheet = SheetNames[i],
      range = "A01:AK80",
      col_names = F )
    
    
    vegetationsdaekke <- select(AllData, 1,3:4)
    vegetationsdaekke <- vegetationsdaekke[10:18,]
    
    colnames(vegetationsdaekke) <- vegetationsdaekke[1,]
    vegetationsdaekke <- vegetationsdaekke[-1, ]
    
    
    #changing '+' to '1'
    vegetationsdaekke$`5m` <- chartr('+', "1", vegetationsdaekke$`5m`)
    vegetationsdaekke$`15m` <- chartr('+', "1", vegetationsdaekke$`15m`)
    
    #adding 'Type' column and changing column names
    vegetationsdaekke$Type <- colnames(vegetationsdaekke)[1]
    names(vegetationsdaekke)[1] <- "Species"
    names(vegetationsdaekke)[2] <- "5m"
    names(vegetationsdaekke)[3] <- "15m"
    
    
    #merging with basic df
    vegetationsdaekkeDF<- merge(x=df, y=vegetationsdaekke, by= NULL)
    
    #merging with Husby df
    husbyDT <- data.frame()
    husbyDT<- rbind(husbyDT, vegetationsdaekkeDF)
    
    print("blomster")
    
    
    
    
    #Importing rows with "græs/urtevegetation"
    GraesUrtevegetation <- select(AllData, 1,3:4)
    GraesUrtevegetation <- GraesUrtevegetation[19:41,]
    
    colnames(GraesUrtevegetation) <- GraesUrtevegetation[1,]
    GraesUrtevegetation <- GraesUrtevegetation[-1, ]
    
    #adding 'Type' column and changing column names
    GraesUrtevegetation$"Type" <- colnames(GraesUrtevegetation)[1]
    names(GraesUrtevegetation)[1] <- "Species"
    names(GraesUrtevegetation)[2] <- "5m"
    names(GraesUrtevegetation)[3] <- "15m"
    
    
    #changing '+' to '1'
    GraesUrtevegetation$`5m` <- chartr('+', "1", GraesUrtevegetation$`5m`)
    GraesUrtevegetation$`15m` <- chartr('+', "1", GraesUrtevegetation$`15m`)
    
    
    #merging with basic df
    GraesUrteDF<- merge(x=df, y=GraesUrtevegetation, by= NULL)
    
    #Merging
    husbyDT<- rbind(husbyDT, GraesUrteDF)
    
    print("graes_urte")
    
    
    
    
    #Importing rows with "substrat"
    substrat <- select(AllData, 1,3:4)
    substrat <- substrat[42:52,]
    
    colnames(substrat) <- substrat[1,]
    substrat <- substrat[-1, ]
    
    #adding 'Type' column and changing column names
    substrat$"Type" <- colnames(substrat)[1]
    names(substrat)[1] <- "Species"
    names(substrat)[2] <- "5m"
    names(substrat)[3] <- "15m"
    
    
    #changing '+' to '1'
    substrat$`5m` <- chartr('+', "1", substrat$`5m`)
    substrat$`15m` <- chartr('+', "1", substrat$`15m`)
    
    
    
    
    #subsetting to the columns we need
    substrat <- subset(substrat, select = c('Species','Type', '5m', '15m'))
    
    #merging with basic df
    substratDF<- merge(x=df, y=substrat, by= NULL)
    
    #Merging
    husbyDT<- rbind(husbyDT, substratDF)
    
    print("substrat")
    
    
    
    
    
    
    #Importing rows with "invasive arter & hydrologi 1"
    invasiveHydrologi1 <- select(AllData, 1,3:4)
    invasiveHydrologi1 <- invasiveHydrologi1[53:63,]
    
    colnames(invasiveHydrologi1) <- invasiveHydrologi1[1,]
    invasiveHydrologi1 <- invasiveHydrologi1[-1, ]
    
    
    #adding 'Type' column and changing column names
    invasiveHydrologi1$"Type" <- colnames(invasiveHydrologi1)[1]
    names(invasiveHydrologi1)[1] <- "Species"
    names(invasiveHydrologi1)[2] <- "5m"
    names(invasiveHydrologi1)[3] <- "15m"
    
    #changing '+' to '1'
    invasiveHydrologi1$`5m` <- chartr('+', "1", invasiveHydrologi1$`5m`)
    invasiveHydrologi1$`15m` <- chartr('+', "1", invasiveHydrologi1$`15m`)
    
    
    #merging with basic df
    invasiveDF<- merge(x=df, y=invasiveHydrologi1, by= NULL)
    
    #Merging
    husbyDT<- rbind(husbyDT, invasiveDF)
    
    print("subvasive")
    
    
    
    
    
    #Importing rows with "Vedplanter - struktur"
    vedplanter <- select(AllData, 5,7:8)
    vedplanter <- vedplanter[10:16,]
    
    colnames(vedplanter) <- vedplanter[1,]
    vedplanter <- vedplanter[-1, ]
    
    
    #adding 'Type' column and changing column names
    vedplanter$"Type" <- colnames(vedplanter)[1]
    names(vedplanter)[1] <- "Species"
    names(vedplanter)[2] <- "5m"
    names(vedplanter)[3] <- "15m"
    
    
    #changing '+' to '1'
    vedplanter$`5m` <- chartr('+', "1", vedplanter$`5m`)
    vedplanter$`15m` <- chartr('+', "1", vedplanter$`15m`)
    
    
    
    #merging with basic df
    vedplanterDF<- merge(x=df, y=vedplanter, by= NULL)
    
    #Merging
    husbyDT<- rbind(husbyDT, vedplanterDF)
    
    print("vedplanter")
    
    
    
    
    
    
    #Importing rows with "Påvirknings- og driftfaktorer"
    paavirkningDrift <- select(AllData, 5,7:8)
    paavirkningDrift <- paavirkningDrift[17:29,]
    
    colnames(paavirkningDrift) <- paavirkningDrift[1,]
    paavirkningDrift <- paavirkningDrift[-1, ]
    
    #adding 'Type' column and changing column names
    paavirkningDrift$"Type" <- colnames(paavirkningDrift)[1]
    names(paavirkningDrift)[1] <- "Species"
    names(paavirkningDrift)[2] <- "5m"
    names(paavirkningDrift)[3] <- "15m"
    
    #changing '+' to '1'
    paavirkningDrift$`5m` <- chartr('+', "1", paavirkningDrift$`5m`)
    paavirkningDrift$`15m` <- chartr('+', "1", paavirkningDrift$`15m`)
    
    
    
    #merging with basic df
    pdDF<- merge(x=df, y=paavirkningDrift, by= NULL)
    
    #Merging
    husbyDT<- rbind(husbyDT, pdDF)
    
    print("paavirkningDrift")
    
    
    
    
    
    
    #Importing rows with "Hydrologi 2"
    hydrologi2 <- select(AllData, 5,7:8)
    hydrologi2 <- hydrologi2[30:32,]
    
    colnames(hydrologi2) <- hydrologi2[1,]
    hydrologi2 <- hydrologi2[-1, ]
    
    #adding 'Type' column and changing column names
    hydrologi2$"Type" <- colnames(hydrologi2)[1]
    names(hydrologi2)[1] <- "Species"
    names(hydrologi2)[2] <- "5m"
    names(hydrologi2)[3] <- "15m"
    
    
    #changing '+' to '1'
    hydrologi2$`5m` <- chartr('+', "1", hydrologi2$`5m` )
    hydrologi2$`15m`  <- chartr('+', "1", hydrologi2$`15m` )
    
    
    
    #merging with basic df
    hydrologiDF<- merge(x=df, y=hydrologi2, by= NULL)
    
    
    #Merging
    husbyDT<- rbind(husbyDT, hydrologiDF)
    
    
    print("hydrologi")
    
    
    
    
    
    
    #Importing rows with "Tegn på græssende dyr"
    animals <- select(AllData, 5,7:8)
    animals <- animals[33:38,]
    
    colnames(animals) <- animals[1,]
    animals <- animals[-1, ]
    
    
    #adding 'Type' column and changing column names
    animals$"Type" <- colnames(animals)[1]
    names(animals)[1] <- "Species"
    names(animals)[2] <- "5m"
    names(animals)[3] <- "15m"
    
    
    #changing '+' to '1'
    animals$`5m` <- chartr('+', "1", animals$`5m`)
    animals$`15m` <- chartr('+', "1", animals$`15m`)
    
    
    #merging with basic df
    animalsDF<- merge(x=df, y=animals, by= NULL)
    
    #Merging
    husbyDT<- rbind(husbyDT, animalsDF)
    
    
    print("animals")
    
    
    
    
    
    
    
    #Importing rows with "Problemarter"
    problem <- select(AllData, 5,7:8)
    problem <- problem[39:45,]
    
    colnames(problem) <- problem[1,]
    problem <- problem[-1, ]
    
    #removing rows containing only NA's
    problem <- problem[rowSums(is.na(problem)) != ncol(problem), ]
    
    
    #adding 'Type' column and changing column names
    problem$"Type" <- colnames(problem)[1]
    names(problem)[1] <- "Species"
    names(problem)[2] <- "5m"
    names(problem)[3] <- "15m"
    
    
    #changing '+' to '1'
    problem$`5m` <- chartr('+', "1", problem$`5m`)
    problem$`15m` <- chartr('+', "1", problem$`15m`)
    
    
    
    
    #merging with basic df
    problemDF<- merge(x=df, y=problem, by= NULL)
    
    #Merging
    husbyDT<- rbind(husbyDT, problemDF)
    
    
    print("problem")
    
    
    
    
    
    
    #pinpoint data
    
    PinpointData <- select(AllData, 10:31)
    PinpointData <- PinpointData[10:67,]
    
    colnames(PinpointData) <- PinpointData[1,]
    PinpointData <- PinpointData[-1, ]
    
    
    #removing rows containing only NA's
    PinpointData <- PinpointData[rowSums(is.na(PinpointData)) != ncol(PinpointData), ]
    
    #removes rows wit only NA values and setting NA's to 0
    #PinpointData <-PinpointData %>% filter(!if_all(1:22, is.na))
    PinpointData$`Vegetationsanalyse - pinpoint`[is.na(PinpointData$`Vegetationsanalyse - pinpoint`)]<-"0"
    
    
    #Adding pinpoint column with zeros as values
    husbyDT[colnames(PinpointData[1][1])] <-0
    
    #creating empty vector for duplicates
    duplicates<- c()
    
    #Updating pinpoint value for duplicates
    for (k in 2:nrow(PinpointData)) {
      for (l in 1:nrow(husbyDT)) {
        if(PinpointData[k,1] == husbyDT[l,5]) {
          PinpointCount<-PinpointData[k,7:22]
          if (!is.character(PinpointCount)==TRUE){
            PinpointCount <- lapply(PinpointCount, as.character)
          }
          PinpointCount[is.na(PinpointCount)]<-"0"
          husbyDT[l,9]=sum(str_count(PinpointCount, "x"))
          duplicates <- append(duplicates, husbyDT[l,5])
        }
      }
    }
    
    #delete duplicate rows in pinpoint data
    PinpointData<-PinpointData[!(PinpointData$`Vegetationsanalyse - pinpoint`== duplicates[1]
                                 | PinpointData$`Vegetationsanalyse - pinpoint`== duplicates[2]
                                 | PinpointData$`Vegetationsanalyse - pinpoint`== duplicates[3]),]
    
    
    #Delete headline "dansk navn"
    PinpointData <- PinpointData[PinpointData$`Vegetationsanalyse - pinpoint` != 'Dansk navn',]
    
    #delete first row (only NA's)
    PinpointData <- PinpointData[-(1),]
    
    #creating count column and for loop that counts x's
    PinpointData$'Count Pinpoint' <-0
    
    for (m in 1:nrow(PinpointData)) {
      PinpointCount<-PinpointData[m,]
      PinpointCount <- PinpointCount[7:ncol(PinpointData)]
      if (!is.character(PinpointCount)==TRUE){
        PinpointCount <- lapply(PinpointCount, as.character)
      }
      PinpointCount[is.na(PinpointCount)]<-"0"
      PinpointData[m,23]=sum(str_count(PinpointCount, "x"))
      
    }
    
    
    #adding 'Type' column and changing column names
    PinpointData$"Type" <- colnames(PinpointData)[1]
    names(PinpointData)[1] <- "Species"
    PinpointData$'5m' <- 0
    PinpointData$'15m' <- 0
    
    
    #subsetting to the columns we need
    PinpointData <- subset(PinpointData, select = c('Species', '5m', '15m','Type', 'Count Pinpoint'))
    
    
    
    #merging with basic df
    PinpointDF<- merge(x=df, y=PinpointData, by= NULL)
    
    #making sure date column is of the date type
    PinpointDF$Dato <- as.character(PinpointDF$Dato)
    
    #renaming columns
    #names(PinpointDF)[9] <- PinpointData[2,2]
    #names(PinpointDF)[8] <- "15m"
    #names(PinpointDF)[7] <- "5m"
    
    names(PinpointDF) <- names(husbyDT)
    
    #Merging
    husbyDT<- rbind(husbyDT, PinpointDF)
    
    print("pinpoint")
    
    
    
    
    #importing rows with habitattype, krondække, vegetationshøjde, fokusart,
    #supplerende vedplanter
    Extra1 <- read_excel(
      path,
      sheet = SheetNames[15],
      range = "A01:AK09",
      col_names = F )
    
    Extra <- select(AllData, 1:37)
    Extra <- Extra[1:9,]
    
    
    #checking that the index for 'habitattype' is correct
    if (is.na(Extra[1,33])==TRUE){
      Extra <- Extra[-1, ]
    }
    
    
    
    #making small datasets
    
    habitattype <- select(Extra, 33:37)
    kronedække <-select(Extra, 5:7)
    kronedække<-kronedække[3:8,]
    vegetationshoejde<-select(Extra, 8:9)
    
    
    m <- matrix("0", ncol = 21, nrow = nrow(husbyDT))
    
    extraData<-data.frame(m)
    
    
    #adding columns and changing column names
    
    
    extraData[1]<-habitattype[3,2]
    extraData[2]<- habitattype[4,2]
    extraData[3]<- habitattype[5,2]
    extraData[4]<-habitattype[4,3]
    extraData[5]<-habitattype[5,3]
    extraData[6]<-habitattype[4,4]
    extraData[7]<-habitattype[5,4]
    extraData[8]<-habitattype[4,5]
    extraData[9]<-habitattype[5,5]
    extraData[10]<- kronedække[3,2]
    extraData[11]<- kronedække[4,2]
    extraData[12]<- kronedække[5,2]
    extraData[13]<- kronedække[6,2]
    extraData[14]<- kronedække[3,3]
    extraData[15]<- kronedække[4,3]
    extraData[16]<- kronedække[5,3]
    extraData[17]<- kronedække[6,3]
    extraData[18]<- vegetationshoejde[4,2]
    extraData[19]<- vegetationshoejde[5,2]
    extraData[20]<- vegetationshoejde[6,2]
    extraData[21]<- vegetationshoejde[7,2]
    
    
    names(extraData)[1] <- "Pinpoint habitattype 1"
    names(extraData)[2] <- "5m habitattype 1"
    names(extraData)[3] <- "15m habitattype 1"
    names(extraData)[4] <- "5m habitattype 2"
    names(extraData)[5] <- "15m habitattype 2"
    names(extraData)[6] <- "5m mosaik 1"
    names(extraData)[7] <- "15m mosaik 1"
    names(extraData)[8] <- "5m mosaik 2"
    names(extraData)[9] <- "15m mosaik 2"
    names(extraData)[10]<- "kronedække 1m nord"
    names(extraData)[11]<- "kronedække 1m oest"
    names(extraData)[12]<- "kronedække 1m syd"
    names(extraData)[13]<- "kronedække 1m vest"
    names(extraData)[14]<- "kronedække 5m nord"
    names(extraData)[15]<- "kronedække 5m oest"
    names(extraData)[16]<- "kronedække 5m syd"
    names(extraData)[17]<- "kronedække 5m vest"
    names(extraData)[18]<- "vegetationshoejde nord"
    names(extraData)[19]<- "vegetationshoejde oest"
    names(extraData)[20]<- "vegetationshoejde syd"
    names(extraData)[21]<- "vegetationshoejde vest"
    
    
    
    #Merging
    
    husbyDT<-cbind(husbyDT, extraData[!names(extraData) %in% names(husbyDT)])
    
    
    
    
    #adding count species column
    
    countSpecies <- matrix(count(unique(husbyDT[5])), ncol = 1, nrow = nrow(husbyDT))
    
    CountSpeciesData<-data.frame(countSpecies)
    
    
    #Merging
    
    husbyDT<-cbind(husbyDT, CountSpeciesData[!names(CountSpeciesData) %in% names(husbyDT)])
    
    
    
    
    
    
    #Changing "æ,ø,å" to "ae,oe,aa"
    husbyDT$Species <- str_replace_all(husbyDT$Species,'å','aa')
    husbyDT$Species <- str_replace_all(husbyDT$Species,'ø','oe')
    husbyDT$Species <- str_replace_all(husbyDT$Species,'æ','ae')
    husbyDT$Species <- str_replace_all(husbyDT$Species,'ë','ë')
    
    
    husbyDT$Type <- str_replace_all(husbyDT$Type,'å','aa')
    husbyDT$Type <- str_replace_all(husbyDT$Type,'ø','oe')
    husbyDT$Type <- str_replace_all(husbyDT$Type,'æ','ae')
    husbyDT$Type <- str_replace_all(husbyDT$Type,'ë','ë')
    
    #Printing Sheetname in case of error
    print(SheetNames[i])
    
    #Creating list of the dataframes
    listOfDFs <- list()
    
    listOfDFs[[i]] <- husbyDT
    
    if (i==1){
      new <- listOfDFs[[1]]
    }
    
    if(i == 2){
      new<-rbind(new,listOfDFs[[2]])
    }
    
    if (i>2){
      new<-rbind(new, listOfDFs[[i]])
    }
    
    
    #Combining the dataframes
    #finalDT <- do.call("rbind", listOfDFs)
    
  }
  
  
  speciesoccured=new[new$Type == "Vegetationsanalyse - pinpoint",]
  speciesoccured1 = speciesoccured[speciesoccured$`Vegetationsanalyse - pinpoint` > 0,]
  
  
  
  plotCount <- count(new[1], new[5] )
  
  new$PlotCount <- as.integer(0)
  
  for (o in 1:nrow(new)){
    if ( any(new[o,5] == plotCount) == TRUE) {
      z <- which(plotCount ==new[o,5])
      new[o,32] <- plotCount[z,2]
    }
  }
  return(new)
}





