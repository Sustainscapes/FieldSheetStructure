#' read_fieldsheet
#'
#'
#' @param path character, the file path of the excel sheet
#' @param SheetNames character vector, sheet names of the excel sheet
#' @param parallel logical, whether to run the function in parallel
#' @param verbose logical, whether to write messages while the function runs, defaults to FALSE
#' @param cores integer, number of cores to use if parallel = TRUE
#' @param type which type of fieldsheet you are using options are Husby or Sinks
#' @return dataframe, a dataframe containing the data from the excel sheet
#' @export

read_fieldsheet <- function(path, SheetNames, type = NULL, parallel = FALSE, cores = NULL, verbose = TRUE){
  if(type == "Husby"){
    Result <- HusbyFieldSheet(path, SheetNames, parallel = FALSE, cores = NULL, verbose = TRUE)
  } else if(type == "Sinks"){
    Result <- SinksFieldSheet(path, SheetNames, parallel = FALSE, cores = NULL, verbose = TRUE)
  }
  return(Result)
}


#' TrojborgFieldSheet
#'
#' This function reads excel files with sheet names and returns a data frame
#'
#' @param path character, the file path of the excel sheet
#' @param SheetNames character vector, sheet names of the excel sheet
#' @param parallel logical, whether to run the function in parallel
#' @param verbose logical, whether to write messages while the function runs, defaults to FALSE
#' @param cores integer, number of cores to use if parallel = TRUE
#' @return dataframe, a dataframe containing the data from the excel sheet
#' @importFrom readxl read_excel
#' @importFrom dplyr select
#' @importFrom dplyr count
#' @importFrom stringr str_count
#' @importFrom stringr str_replace_all
#' @importFrom foreach %dopar% foreach
#' @importFrom doParallel registerDoParallel
#' @importFrom parallel makeCluster
#' @importFrom parallel stopCluster
#' @export


TrojborgFieldSheet <- function(path, SheetNames, parallel = FALSE, cores = NULL, verbose = TRUE){
  #SheetNames <- excel_sheets(path)
  Species <- NULL
  lengthLoop <- length(SheetNames)
  if(!parallel){
    for (i in 1:lengthLoop) {
      if(verbose){
        message(paste("starting", SheetNames[i]))
      }

      PlotDateInventorLocation <- read_excel(path, sheet = SheetNames[i],
                                             skip = 0, n_max = 1, col_names = F)
      PlotDateInventorLocation[is.na(PlotDateInventorLocation)] <- 0
      col_names <- PlotDateInventorLocation[seq(1, length(PlotDateInventorLocation),
                                                2)]
      PlotDateInventorLocation <- PlotDateInventorLocation[1,
      ]
      if (PlotDateInventorLocation[8] == 0) {
        PlotDateInventorLocation[8] <- "missing"
      }
      df <- col_names[1:4]
      colnames(df) <- col_names[1:4]
      if (SheetNames[i]=="T_E_C2_037_A"){
        df$Dato <- PlotDateInventorLocation[[4]]
      } else{
        df$Dato <- as.Date(df$Dato, format = "%Y.%m.%d")
      }

      for (j in 1:4) {
        if (PlotDateInventorLocation[4] == 0) {
          df$Dato <- NA
          df[1, 1] <- PlotDateInventorLocation[1 * 2]
          df[1, 2] <- NA
          df[1, 3] <- PlotDateInventorLocation[3 * 2]
          df[1, 4] <- PlotDateInventorLocation[4 * 2]
        } else if (SheetNames[i]=="T_E_C2_037_A") {
          df[1, 1] <- PlotDateInventorLocation[1 * 2]
          df[1, 3] <- PlotDateInventorLocation[3 * 2]
          df[1, 4] <- PlotDateInventorLocation[4 * 2]
        } else {
          df[1, j] <- PlotDateInventorLocation[j * 2]
        }
      }
      AllData <- read_excel(path, sheet = SheetNames[i], range = "A01:AK80",
                            col_names = F)
      vegetationsdaekke <- select(AllData, 1, 3:4)
      vegetationsdaekke <- vegetationsdaekke[10:18, ]
      colnames(vegetationsdaekke) <- vegetationsdaekke[1, ]
      vegetationsdaekke <- vegetationsdaekke[-1, ]
      vegetationsdaekke$`5m` <- chartr("+", "1", vegetationsdaekke$`5m`)
      vegetationsdaekke$`15m` <- chartr("+", "1", vegetationsdaekke$`15m`)
      vegetationsdaekke$Type <- colnames(vegetationsdaekke)[1]
      names(vegetationsdaekke)[1] <- "Species"
      names(vegetationsdaekke)[2] <- "5m"
      names(vegetationsdaekke)[3] <- "15m"
      vegetationsdaekkeDF <- merge(x = df, y = vegetationsdaekke,
                                   by = NULL)
      TrojborgDT <- data.frame()
      TrojborgDT <- rbind(TrojborgDT, vegetationsdaekkeDF)
      if(verbose){
        print("blomster")
      }
      GraesUrtevegetation <- select(AllData, 1, 3:4)
      GraesUrtevegetation <- GraesUrtevegetation[19:41, ]
      colnames(GraesUrtevegetation) <- GraesUrtevegetation[1,
      ]
      GraesUrtevegetation <- GraesUrtevegetation[-1, ]
      GraesUrtevegetation$Type <- colnames(GraesUrtevegetation)[1]
      names(GraesUrtevegetation)[1] <- "Species"
      names(GraesUrtevegetation)[2] <- "5m"
      names(GraesUrtevegetation)[3] <- "15m"
      GraesUrtevegetation$`5m` <- chartr("+", "1", GraesUrtevegetation$`5m`)
      GraesUrtevegetation$`15m` <- chartr("+", "1", GraesUrtevegetation$`15m`)
      GraesUrteDF <- merge(x = df, y = GraesUrtevegetation,
                           by = NULL)
      TrojborgDT <- rbind(TrojborgDT, GraesUrteDF)
      if(verbose){
        print("graes_urte")
      }
      substrat <- select(AllData, 1, 3:4)
      substrat <- substrat[42:52, ]
      colnames(substrat) <- substrat[1, ]
      substrat <- substrat[-1, ]
      substrat$Type <- colnames(substrat)[1]
      names(substrat)[1] <- "Species"
      names(substrat)[2] <- "5m"
      names(substrat)[3] <- "15m"
      substrat$`5m` <- chartr("+", "1", substrat$`5m`)
      substrat$`15m` <- chartr("+", "1", substrat$`15m`)
      substrat <- subset(substrat, select = c("Species", "Type",
                                              "5m", "15m"))
      substratDF <- merge(x = df, y = substrat, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, substratDF)
      if(verbose){
        print("substrat")
      }
      invasiveHydrologi1 <- select(AllData, 1, 3:4)
      invasiveHydrologi1 <- invasiveHydrologi1[53:63, ]
      colnames(invasiveHydrologi1) <- invasiveHydrologi1[1,
      ]
      invasiveHydrologi1 <- invasiveHydrologi1[-1, ]
      invasiveHydrologi1$Type <- colnames(invasiveHydrologi1)[1]
      names(invasiveHydrologi1)[1] <- "Species"
      names(invasiveHydrologi1)[2] <- "5m"
      names(invasiveHydrologi1)[3] <- "15m"
      invasiveHydrologi1$`5m` <- chartr("+", "1", invasiveHydrologi1$`5m`)
      invasiveHydrologi1$`15m` <- chartr("+", "1", invasiveHydrologi1$`15m`)
      invasiveDF <- merge(x = df, y = invasiveHydrologi1, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, invasiveDF)
      if(verbose){
        print("subvasive")
      }
      vedplanter <- select(AllData, 5, 7:8)
      vedplanter <- vedplanter[10:16, ]
      colnames(vedplanter) <- vedplanter[1, ]
      vedplanter <- vedplanter[-1, ]
      vedplanter$Type <- colnames(vedplanter)[1]
      names(vedplanter)[1] <- "Species"
      names(vedplanter)[2] <- "5m"
      names(vedplanter)[3] <- "15m"
      vedplanter$`5m` <- chartr("+", "1", vedplanter$`5m`)
      vedplanter$`15m` <- chartr("+", "1", vedplanter$`15m`)
      vedplanterDF <- merge(x = df, y = vedplanter, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, vedplanterDF)
      if(verbose){
        print("vedplanter")
      }

      paavirkningDrift <- select(AllData, 5, 7:8)
      paavirkningDrift <- paavirkningDrift[17:29, ]
      colnames(paavirkningDrift) <- paavirkningDrift[1, ]
      paavirkningDrift <- paavirkningDrift[-1, ]
      paavirkningDrift$Type <- colnames(paavirkningDrift)[1]
      names(paavirkningDrift)[1] <- "Species"
      names(paavirkningDrift)[2] <- "5m"
      names(paavirkningDrift)[3] <- "15m"
      paavirkningDrift$`5m` <- chartr("+", "1", paavirkningDrift$`5m`)
      paavirkningDrift$`15m` <- chartr("+", "1", paavirkningDrift$`15m`)
      pdDF <- merge(x = df, y = paavirkningDrift, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, pdDF)
      if(verbose){
        print("paavirkningDrift")
      }
      hydrologi2 <- select(AllData, 5, 7:8)
      hydrologi2 <- hydrologi2[30:32, ]
      colnames(hydrologi2) <- hydrologi2[1, ]
      hydrologi2 <- hydrologi2[-1, ]
      hydrologi2$Type <- colnames(hydrologi2)[1]
      names(hydrologi2)[1] <- "Species"
      names(hydrologi2)[2] <- "5m"
      names(hydrologi2)[3] <- "15m"
      hydrologi2$`5m` <- chartr("+", "1", hydrologi2$`5m`)
      hydrologi2$`15m` <- chartr("+", "1", hydrologi2$`15m`)
      hydrologiDF <- merge(x = df, y = hydrologi2, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, hydrologiDF)
      if(verbose){
        print("hydrologi")
      }
      animals <- select(AllData, 5, 7:8)
      animals <- animals[33:38, ]
      colnames(animals) <- animals[1, ]
      animals <- animals[-1, ]
      animals$Type <- colnames(animals)[1]
      names(animals)[1] <- "Species"
      names(animals)[2] <- "5m"
      names(animals)[3] <- "15m"
      animals$`5m` <- chartr("+", "1", animals$`5m`)
      animals$`15m` <- chartr("+", "1", animals$`15m`)
      animalsDF <- merge(x = df, y = animals, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, animalsDF)
      if(verbose){
        print("animals")
      }
      problem <- select(AllData, 5, 7:8)
      problem <- problem[39:45, ]
      colnames(problem) <- problem[1, ]
      problem <- problem[-1, ]
      problem <- problem[rowSums(is.na(problem)) != ncol(problem),
      ]
      problem$Type <- colnames(problem)[1]
      names(problem)[1] <- "Species"
      names(problem)[2] <- "5m"
      names(problem)[3] <- "15m"
      problem$`5m` <- chartr("+", "1", problem$`5m`)
      problem$`15m` <- chartr("+", "1", problem$`15m`)
      problemDF <- merge(x = df, y = problem, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, problemDF)
      if(verbose){
        print("problem")
      }

      PinpointData <- select(AllData, 10:31)
      PinpointData <- PinpointData[10:67, ]
      colnames(PinpointData) <- PinpointData[1, ]
      PinpointData <- PinpointData[-1, ]
      PinpointData <- PinpointData[rowSums(is.na(PinpointData)) !=
                                     ncol(PinpointData), ]
      PinpointData$`Vegetationsanalyse - pinpoint`[is.na(PinpointData$`Vegetationsanalyse - pinpoint`)] <- "0"
      TrojborgDT[colnames(PinpointData[1][1])] <- 0
      duplicates <- c()
      for (k in 2:nrow(PinpointData)) {
        for (l in 1:nrow(TrojborgDT)) {
          if (PinpointData[[k, 1]] == TrojborgDT[l, 5]) {
            PinpointCount <- PinpointData[k, 7:22]
            if (!is.character(PinpointCount) == TRUE) {
              print("count error")
              PinpointCount <- lapply(PinpointCount, as.character)
            }
            PinpointCount[is.na(PinpointCount)] <- "0"
            TrojborgDT[l, 9] = sum(str_count(PinpointCount,
                                          "x"))
            duplicates <- append(duplicates, TrojborgDT[l,
                                                     5])
          }
        }
      }
      PinpointData <- PinpointData[!(PinpointData$`Vegetationsanalyse - pinpoint` ==
                                       duplicates[1] | PinpointData$`Vegetationsanalyse - pinpoint` ==
                                       duplicates[2] | PinpointData$`Vegetationsanalyse - pinpoint` ==
                                       duplicates[3]), ]
      PinpointData <- PinpointData[PinpointData$`Vegetationsanalyse - pinpoint` !=
                                     "Dansk navn", ]
      PinpointData <- PinpointData[-(1), ]
      PinpointData$"Count Pinpoint" <- 0
      for (m in 1:nrow(PinpointData)) {
        PinpointCount <- PinpointData[m, ]
        PinpointCount <- PinpointCount[7:ncol(PinpointData)]
        if (!is.character(PinpointCount) == TRUE) {
          PinpointCount <- lapply(PinpointCount, as.character)
        }
        PinpointCount[is.na(PinpointCount)] <- "0"
        PinpointData[m, 23] = sum(str_count(PinpointCount,
                                            "x"))
      }
      PinpointData$Type <- colnames(PinpointData)[1]
      names(PinpointData)[1] <- "Species"
      PinpointData$"5m" <- 0
      PinpointData$"15m" <- 0
      PinpointData <- subset(PinpointData, select = c("Species",
                                                      "5m", "15m", "Type", "Count Pinpoint"))

      PinpointDF <- merge(x = df, y = PinpointData, by = NULL)
      if (SheetNames[i]=="T_E_C2_037_A"){
        df$Dato <- PlotDateInventorLocation[[4]]
      } else{
        df$Dato <- as.Date(df$Dato, format = "%Y.%m.%d")
      }
      #PinpointDF$Dato <- as.character(PinpointDF$Dato)
      names(PinpointDF) <- names(TrojborgDT)
      PinpointDF <- PinpointDF[1:9]
      TrojborgDT <- rbind(TrojborgDT, PinpointDF)
      if(verbose){
        print("pinpoint")
      }
      Extra1 <- read_excel(path, sheet = SheetNames[15], range = "A01:AK09",
                           col_names = F)
      Extra <- select(AllData, 1:37)
      Extra <- Extra[1:9, ]
      if (is.na(Extra[1, 33]) == TRUE) {
        Extra <- Extra[-1, ]
      }
      habitattype <- select(Extra, 33:37)
      kronedaekke <- select(Extra, 5:7)
      kronedaekke <- kronedaekke[3:8, ]
      vegetationshoejde <- select(Extra, 8:9)
      m <- matrix("0", ncol = 21, nrow = nrow(TrojborgDT))
      extraData <- data.frame(m)
      extraData[1] <- habitattype[3, 2]
      extraData[2] <- habitattype[4, 2]
      extraData[3] <- habitattype[5, 2]
      extraData[4] <- habitattype[4, 3]
      extraData[5] <- habitattype[5, 3]
      extraData[6] <- habitattype[4, 4]
      extraData[7] <- habitattype[5, 4]
      extraData[8] <- habitattype[4, 5]
      extraData[9] <- habitattype[5, 5]
      extraData[10] <- kronedaekke[3, 2]
      extraData[11] <- kronedaekke[4, 2]
      extraData[12] <- kronedaekke[5, 2]
      extraData[13] <- kronedaekke[6, 2]
      extraData[14] <- kronedaekke[3, 3]
      extraData[15] <- kronedaekke[4, 3]
      extraData[16] <- kronedaekke[5, 3]
      extraData[17] <- kronedaekke[6, 3]
      extraData[18] <- vegetationshoejde[4, 2]
      extraData[19] <- vegetationshoejde[5, 2]
      extraData[20] <- vegetationshoejde[6, 2]
      extraData[21] <- vegetationshoejde[7, 2]
      names(extraData)[1] <- "Pinpoint habitattype 1"
      names(extraData)[2] <- "5m habitattype 1"
      names(extraData)[3] <- "15m habitattype 1"
      names(extraData)[4] <- "5m habitattype 2"
      names(extraData)[5] <- "15m habitattype 2"
      names(extraData)[6] <- "5m mosaik 1"
      names(extraData)[7] <- "15m mosaik 1"
      names(extraData)[8] <- "5m mosaik 2"
      names(extraData)[9] <- "15m mosaik 2"
      names(extraData)[10] <- "kronedaekke 1m nord"
      names(extraData)[11] <- "kronedaekke 1m oest"
      names(extraData)[12] <- "kronedaekke 1m syd"
      names(extraData)[13] <- "kronedaekke 1m vest"
      names(extraData)[14] <- "kronedaekke 5m nord"
      names(extraData)[15] <- "kronedaekke 5m oest"
      names(extraData)[16] <- "kronedaekke 5m syd"
      names(extraData)[17] <- "kronedaekke 5m vest"
      names(extraData)[18] <- "vegetationshoejde nord"
      names(extraData)[19] <- "vegetationshoejde oest"
      names(extraData)[20] <- "vegetationshoejde syd"
      names(extraData)[21] <- "vegetationshoejde vest"
      TrojborgDT <- cbind(TrojborgDT, extraData[!names(extraData) %in%
                                            names(TrojborgDT)])
      countSpecies <- matrix(count(unique(TrojborgDT[5])), ncol = 1,
                             nrow = nrow(TrojborgDT))
      CountSpeciesData <- data.frame(countSpecies)
      TrojborgDT <- cbind(TrojborgDT, CountSpeciesData[!names(CountSpeciesData) %in%
                                                   names(TrojborgDT)])
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00e5", "aa")
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00f8", "oe")
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00e6", "ae")
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00eb",
                                                  "e")
      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00e5", "aa")
      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00f8", "oe")
      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00e6", "ae")
      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00eb", "e")
      if(verbose){
        print(SheetNames[i])
      }
      listOfDFs <- list()
      listOfDFs[[i]] <- TrojborgDT
      if (i == 1) {
        new <- listOfDFs[[1]]
      }
      if (i == 2) {
        new <- rbind(new, listOfDFs[[2]])
      }
      if (i > 2) {
        new <- rbind(new, listOfDFs[[i]])
      }
      if(verbose){
        message(paste(i, "of" , length(SheetNames), "ready", Sys.time()))
      }
    }
  } else if(parallel){
    cl <- makeCluster(cores)
    registerDoParallel(cl)
    new <- foreach(i=1:lengthLoop, .errorhandling = "pass", .packages = c("readxl", "dplyr", "stringr"), .combine = rbind) %dopar%{
      PlotDateInventorLocation <- read_excel(path, sheet = SheetNames[i],
                                             skip = 0, n_max = 1, col_names = F)
      PlotDateInventorLocation[is.na(PlotDateInventorLocation)] <- 0
      col_names <- PlotDateInventorLocation[seq(1, length(PlotDateInventorLocation),
                                                2)]
      PlotDateInventorLocation <- PlotDateInventorLocation[1,
      ]
      if (PlotDateInventorLocation[8] == 0) {
        PlotDateInventorLocation[8] <- "missing"
      }
      df <- col_names[1:4]
      colnames(df) <- col_names[1:4]
      df$Dato <- as.Date(df$Dato, format = "%Y.%m.%d")
      for (j in 1:4) {
        if (PlotDateInventorLocation[4] == 0) {
          df$Dato <- NA
          df[1, 1] <- PlotDateInventorLocation[1 * 2]
          df[1, 2] <- NA
          df[1, 3] <- PlotDateInventorLocation[3 * 2]
          df[1, 4] <- PlotDateInventorLocation[4 * 2]
        }
        else {
          df[1, j] <- PlotDateInventorLocation[j * 2]
        }
      }
      AllData <- read_excel(path, sheet = SheetNames[i], range = "A01:AK80",
                            col_names = F)
      vegetationsdaekke <- select(AllData, 1, 3:4)
      vegetationsdaekke <- vegetationsdaekke[10:18, ]
      colnames(vegetationsdaekke) <- vegetationsdaekke[1, ]
      vegetationsdaekke <- vegetationsdaekke[-1, ]
      vegetationsdaekke$`5m` <- chartr("+", "1", vegetationsdaekke$`5m`)
      vegetationsdaekke$`15m` <- chartr("+", "1", vegetationsdaekke$`15m`)
      vegetationsdaekke$Type <- colnames(vegetationsdaekke)[1]
      names(vegetationsdaekke)[1] <- "Species"
      names(vegetationsdaekke)[2] <- "5m"
      names(vegetationsdaekke)[3] <- "15m"
      vegetationsdaekkeDF <- merge(x = df, y = vegetationsdaekke,
                                   by = NULL)
      TrojborgDT <- data.frame()
      TrojborgDT <- rbind(TrojborgDT, vegetationsdaekkeDF)
      GraesUrtevegetation <- select(AllData, 1, 3:4)
      GraesUrtevegetation <- GraesUrtevegetation[19:41, ]
      colnames(GraesUrtevegetation) <- GraesUrtevegetation[1,
      ]
      GraesUrtevegetation <- GraesUrtevegetation[-1, ]
      GraesUrtevegetation$Type <- colnames(GraesUrtevegetation)[1]
      names(GraesUrtevegetation)[1] <- "Species"
      names(GraesUrtevegetation)[2] <- "5m"
      names(GraesUrtevegetation)[3] <- "15m"
      GraesUrtevegetation$`5m` <- chartr("+", "1", GraesUrtevegetation$`5m`)
      GraesUrtevegetation$`15m` <- chartr("+", "1", GraesUrtevegetation$`15m`)
      GraesUrteDF <- merge(x = df, y = GraesUrtevegetation,
                           by = NULL)
      TrojborgDT <- rbind(TrojborgDT, GraesUrteDF)
      substrat <- select(AllData, 1, 3:4)
      substrat <- substrat[42:52, ]
      colnames(substrat) <- substrat[1, ]
      substrat <- substrat[-1, ]
      substrat$Type <- colnames(substrat)[1]
      names(substrat)[1] <- "Species"
      names(substrat)[2] <- "5m"
      names(substrat)[3] <- "15m"
      substrat$`5m` <- chartr("+", "1", substrat$`5m`)
      substrat$`15m` <- chartr("+", "1", substrat$`15m`)
      substrat <- subset(substrat, select = c("Species", "Type",
                                              "5m", "15m"))
      substratDF <- merge(x = df, y = substrat, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, substratDF)
      invasiveHydrologi1 <- select(AllData, 1, 3:4)
      invasiveHydrologi1 <- invasiveHydrologi1[53:63, ]
      colnames(invasiveHydrologi1) <- invasiveHydrologi1[1,
      ]
      invasiveHydrologi1 <- invasiveHydrologi1[-1, ]
      invasiveHydrologi1$Type <- colnames(invasiveHydrologi1)[1]
      names(invasiveHydrologi1)[1] <- "Species"
      names(invasiveHydrologi1)[2] <- "5m"
      names(invasiveHydrologi1)[3] <- "15m"
      invasiveHydrologi1$`5m` <- chartr("+", "1", invasiveHydrologi1$`5m`)
      invasiveHydrologi1$`15m` <- chartr("+", "1", invasiveHydrologi1$`15m`)
      invasiveDF <- merge(x = df, y = invasiveHydrologi1, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, invasiveDF)
      vedplanter <- select(AllData, 5, 7:8)
      vedplanter <- vedplanter[10:16, ]
      colnames(vedplanter) <- vedplanter[1, ]
      vedplanter <- vedplanter[-1, ]
      vedplanter$Type <- colnames(vedplanter)[1]
      names(vedplanter)[1] <- "Species"
      names(vedplanter)[2] <- "5m"
      names(vedplanter)[3] <- "15m"
      vedplanter$`5m` <- chartr("+", "1", vedplanter$`5m`)
      vedplanter$`15m` <- chartr("+", "1", vedplanter$`15m`)
      vedplanterDF <- merge(x = df, y = vedplanter, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, vedplanterDF)
      paavirkningDrift <- select(AllData, 5, 7:8)
      paavirkningDrift <- paavirkningDrift[17:29, ]
      colnames(paavirkningDrift) <- paavirkningDrift[1, ]
      paavirkningDrift <- paavirkningDrift[-1, ]
      paavirkningDrift$Type <- colnames(paavirkningDrift)[1]
      names(paavirkningDrift)[1] <- "Species"
      names(paavirkningDrift)[2] <- "5m"
      names(paavirkningDrift)[3] <- "15m"
      paavirkningDrift$`5m` <- chartr("+", "1", paavirkningDrift$`5m`)
      paavirkningDrift$`15m` <- chartr("+", "1", paavirkningDrift$`15m`)
      pdDF <- merge(x = df, y = paavirkningDrift, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, pdDF)
      hydrologi2 <- select(AllData, 5, 7:8)
      hydrologi2 <- hydrologi2[30:32, ]
      colnames(hydrologi2) <- hydrologi2[1, ]
      hydrologi2 <- hydrologi2[-1, ]
      hydrologi2$Type <- colnames(hydrologi2)[1]
      names(hydrologi2)[1] <- "Species"
      names(hydrologi2)[2] <- "5m"
      names(hydrologi2)[3] <- "15m"
      hydrologi2$`5m` <- chartr("+", "1", hydrologi2$`5m`)
      hydrologi2$`15m` <- chartr("+", "1", hydrologi2$`15m`)
      hydrologiDF <- merge(x = df, y = hydrologi2, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, hydrologiDF)
      animals <- select(AllData, 5, 7:8)
      animals <- animals[33:38, ]
      colnames(animals) <- animals[1, ]
      animals <- animals[-1, ]
      animals$Type <- colnames(animals)[1]
      names(animals)[1] <- "Species"
      names(animals)[2] <- "5m"
      names(animals)[3] <- "15m"
      animals$`5m` <- chartr("+", "1", animals$`5m`)
      animals$`15m` <- chartr("+", "1", animals$`15m`)
      animalsDF <- merge(x = df, y = animals, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, animalsDF)
      problem <- select(AllData, 5, 7:8)
      problem <- problem[39:45, ]
      colnames(problem) <- problem[1, ]
      problem <- problem[-1, ]
      problem <- problem[rowSums(is.na(problem)) != ncol(problem),
      ]
      problem$Type <- colnames(problem)[1]
      names(problem)[1] <- "Species"
      names(problem)[2] <- "5m"
      names(problem)[3] <- "15m"
      problem$`5m` <- chartr("+", "1", problem$`5m`)
      problem$`15m` <- chartr("+", "1", problem$`15m`)
      problemDF <- merge(x = df, y = problem, by = NULL)
      TrojborgDT <- rbind(TrojborgDT, problemDF)
      PinpointData <- select(AllData, 10:31)
      PinpointData <- PinpointData[10:67, ]
      colnames(PinpointData) <- PinpointData[1, ]
      PinpointData <- PinpointData[-1, ]
      PinpointData <- PinpointData[rowSums(is.na(PinpointData)) !=
                                     ncol(PinpointData), ]
      PinpointData$`Vegetationsanalyse - pinpoint`[is.na(PinpointData$`Vegetationsanalyse - pinpoint`)] <- "0"
      TrojborgDT[colnames(PinpointData[1][1])] <- 0
      duplicates <- c()
      for (k in 2:nrow(PinpointData)) {
        for (l in 1:nrow(TrojborgDT)) {
          if (PinpointData[[k, 1]] == TrojborgDT[l, 5]) {
            PinpointCount <- PinpointData[k, 7:22]
            if (!is.character(PinpointCount) == TRUE) {
              PinpointCount <- lapply(PinpointCount, as.character)
            }
            PinpointCount[is.na(PinpointCount)] <- "0"
            TrojborgDT[l, 9] = sum(str_count(PinpointCount,
                                          "x"))
            duplicates <- append(duplicates, TrojborgDT[l,
                                                     5])
          }
          print("pinpoint")
        }
      }
      PinpointData <- PinpointData[!(PinpointData$`Vegetationsanalyse - pinpoint` ==
                                       duplicates[1] | PinpointData$`Vegetationsanalyse - pinpoint` ==
                                       duplicates[2] | PinpointData$`Vegetationsanalyse - pinpoint` ==
                                       duplicates[3]), ]
      PinpointData <- PinpointData[PinpointData$`Vegetationsanalyse - pinpoint` !=
                                     "Dansk navn", ]
      PinpointData <- PinpointData[-(1), ]
      PinpointData$"Count Pinpoint" <- 0
      for (m in 1:nrow(PinpointData)) {
        PinpointCount <- PinpointData[m, ]
        PinpointCount <- PinpointCount[7:ncol(PinpointData)]
        if (!is.character(PinpointCount) == TRUE) {
          PinpointCount <- lapply(PinpointCount, as.character)
        }
        PinpointCount[is.na(PinpointCount)] <- "0"
        PinpointData[m, 23] = sum(str_count(PinpointCount,
                                            "x"))
      }
      PinpointData$Type <- colnames(PinpointData)[1]
      names(PinpointData)[1] <- "Species"
      PinpointData$"5m" <- 0
      PinpointData$"15m" <- 0
      PinpointData <- subset(PinpointData, select = c("Species",
                                                      "5m", "15m", "Type", "Count Pinpoint"))
      PinpointDF <- merge(x = df, y = PinpointData, by = NULL)
      PinpointDF$Dato <- as.character(PinpointDF$Dato)
      names(PinpointDF) <- names(TrojborgDT)
      TrojborgDT <- rbind(TrojborgDT, PinpointDF)
      Extra1 <- read_excel(path, sheet = SheetNames[15], range = "A01:AK09",
                           col_names = F)
      Extra <- select(AllData, 1:37)
      Extra <- Extra[1:9, ]
      if (is.na(Extra[1, 33]) == TRUE) {
        Extra <- Extra[-1, ]
      }
      habitattype <- select(Extra, 33:37)
      kronedaekke <- select(Extra, 5:7)
      kronedaekke <- kronedaekke[3:8, ]
      vegetationshoejde <- select(Extra, 8:9)
      m <- matrix("0", ncol = 21, nrow = nrow(TrojborgDT))
      extraData <- data.frame(m)
      extraData[1] <- habitattype[3, 2]
      extraData[2] <- habitattype[4, 2]
      extraData[3] <- habitattype[5, 2]
      extraData[4] <- habitattype[4, 3]
      extraData[5] <- habitattype[5, 3]
      extraData[6] <- habitattype[4, 4]
      extraData[7] <- habitattype[5, 4]
      extraData[8] <- habitattype[4, 5]
      extraData[9] <- habitattype[5, 5]
      extraData[10] <- kronedaekke[3, 2]
      extraData[11] <- kronedaekke[4, 2]
      extraData[12] <- kronedaekke[5, 2]
      extraData[13] <- kronedaekke[6, 2]
      extraData[14] <- kronedaekke[3, 3]
      extraData[15] <- kronedaekke[4, 3]
      extraData[16] <- kronedaekke[5, 3]
      extraData[17] <- kronedaekke[6, 3]
      extraData[18] <- vegetationshoejde[4, 2]
      extraData[19] <- vegetationshoejde[5, 2]
      extraData[20] <- vegetationshoejde[6, 2]
      extraData[21] <- vegetationshoejde[7, 2]
      names(extraData)[1] <- "Pinpoint habitattype 1"
      names(extraData)[2] <- "5m habitattype 1"
      names(extraData)[3] <- "15m habitattype 1"
      names(extraData)[4] <- "5m habitattype 2"
      names(extraData)[5] <- "15m habitattype 2"
      names(extraData)[6] <- "5m mosaik 1"
      names(extraData)[7] <- "15m mosaik 1"
      names(extraData)[8] <- "5m mosaik 2"
      names(extraData)[9] <- "15m mosaik 2"
      names(extraData)[10] <- "kronedaekke 1m nord"
      names(extraData)[11] <- "kronedaekke 1m oest"
      names(extraData)[12] <- "kronedaekke 1m syd"
      names(extraData)[13] <- "kronedaekke 1m vest"
      names(extraData)[14] <- "kronedaekke 5m nord"
      names(extraData)[15] <- "kronedaekke 5m oest"
      names(extraData)[16] <- "kronedaekke 5m syd"
      names(extraData)[17] <- "kronedaekke 5m vest"
      names(extraData)[18] <- "vegetationshoejde nord"
      names(extraData)[19] <- "vegetationshoejde oest"
      names(extraData)[20] <- "vegetationshoejde syd"
      names(extraData)[21] <- "vegetationshoejde vest"
      TrojborgDT <- cbind(TrojborgDT, extraData[!names(extraData) %in%
                                            names(TrojborgDT)])
      countSpecies <- matrix(count(unique(TrojborgDT[5])), ncol = 1,
                             nrow = nrow(TrojborgDT))
      CountSpeciesData <- data.frame(countSpecies)
      TrojborgDT <- cbind(TrojborgDT, CountSpeciesData[!names(CountSpeciesData) %in%
                                                   names(TrojborgDT)])
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00e5", "aa")
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00f8", "oe")
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00e6", "ae")
      TrojborgDT$Species <- stringr::str_replace_all(TrojborgDT$Species, "\\u00eb", "e")


      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00e5", "aa")
      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00f8", "oe")
      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00e6", "ae")
      TrojborgDT$Type <- stringr::str_replace_all(TrojborgDT$Type, "\\u00eb", "e")
      TrojborgDT
    }

    stopCluster(cl)

  }

  speciesoccured = new[new$Type == "Vegetationsanalyse - pinpoint",
  ]
  speciesoccured1 = speciesoccured[speciesoccured$`Vegetationsanalyse - pinpoint` >
                                     0, ]
  new <- dplyr::filter(new, !is.na(Species))
  plotCount <- count(new[1], new[5])
  new$PlotCount <- as.integer(0)

  for (o in 1:nrow(new)) {
    if (any(new[o, 5] == plotCount) == TRUE) {
      z <- which(plotCount == new[o, 5])
      new[o, 32] <- plotCount[z, 2]
    }
  }

  return(new)
}

SheetNames = excel_sheets("/Users/heidilunde/Documents/SustainScapes/DataStructuring/FieldSheetsTrojborg_df.xlsm")
SheetNames = SheetNames[1:88]
tester <- TrojborgFieldSheet("/Users/heidilunde/Documents/SustainScapes/DataStructuring/FieldSheetsTrojborg_df.xlsm", SheetNames)
