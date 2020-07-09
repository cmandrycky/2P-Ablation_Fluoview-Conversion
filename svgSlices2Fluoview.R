svgSlices2Fluoview <- function(zStep = 1, numRepeats = 1, startingZ = 0){
  library(XML)
  library(dplyr)
  library(tidyverse)
  library(dplyr)
  library(RODBC)
  
  polygonEstimate <- function(svgDir){
    filenames <- paste(svgDir,"\\", list.files(path = svgDir,pattern = ".svg"), sep="")
    
    counter <- 0
    for (file in filenames) {
      svg <- readChar(file, file.info(file)$size)
      doc <- htmlParse(svg)
      p <- xpathSApply(doc, "//polygon", xmlGetAttr, "points")
      counter <- counter + length(p)
    }
    return(counter)
  }
  
  GetXY <- function(listIndex, listValue) {
    #Conversion factor for microscope microns to pixels
    micro2Pix <- 2.01574803
    
    data.frame(listValue) %>%
      rename(X = X1, Y = X2) %>%
      mutate(ShapeNum = listIndex,
             X = floor(X*micro2Pix),
             Y = floor(Y*micro2Pix)) %>%
      select(ShapeNum, X, Y)
  }
  
  
  frameTimeEstimate <- function(w,h){
    #Returns estimated frame time in (ms)
    frame <- 20.416 + numRepeats*h*(w*.00197712+1.10997349)
    if(frame < 100){
      return(100)
    } else if(frame > 3300) {
      return(3300)
    } else {
      return(round(frame*1.10))
    }
  }
  
  #Make a copy of the original db file and place it in user defined directory. Name will be based on svg directory file name
  defaultDBpath <- c("C:\\PathToDefaultDB")
  
  svgDirectory <- choose.dir(caption = "Select SVG slices folder")
  
  numPolygons <- polygonEstimate(svgDirectory)
  if(numPolygons > 2999){
    stop(paste("Too many polygons for single protocol file. Break into smaller segments:",numPolygons))
  } else{
    print(paste("Number of polygons:", numPolygons))
  }
  
  #second line names the new DB file after the svgDirectory
  dbSaveToPath <- paste(choose.dir(caption = "Where do you want to save the DB file to?"),
                        "\\",sapply(strsplit(svgDirectory, "\\\\"), tail, 1),".mdb",
                        sep = "")
  
  file.copy(defaultDBpath,
            dbSaveToPath,
            copy.mode = TRUE,
            copy.date = FALSE)
  
  filenames <- paste(svgDirectory,"\\", list.files(path = svgDirectory,pattern = ".svg"), sep="")
  
  
  rpCounter = 2 #first TaskID is a prototype
  zStepSize = zStep*1000 #micron
  zHeight = startingZ*1000 #micron
  
  dbName <- paste("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=",dbSaveToPath,sep = "")
  db <- odbcDriverConnect(dbName)
  
  #editing
  taskClip <- sqlQuery( db , paste("select * from IMAGING_TASK_CLIP"))
  taskInfo <- sqlQuery( db , paste("select * from IMAGING_TASK_INFO"))
  
  #duplicating
  channelInfo <- sqlQuery( db , paste("select * from IMAGING_CHANNEL_INFO"))
  taskLaser <- sqlQuery( db , paste("select * from IMAGING_TASK_LASER"))
  taskMatlinfo <- sqlQuery( db , paste("select * from IMAGING_TASK_MATLINFO"))
  taskScaninfo <- sqlQuery( db , paste("select * from IMAGING_TASK_SCANINFO"))
  taskScanrange <- sqlQuery( db , paste("select * from IMAGING_TASK_SCANRANGE"))
  
  #different spelling
  taskBar <- sqlQuery( db , paste("select * from IMAGING_TASK_BAR"))
  
  
  for (file in filenames) {
    svg <- readChar(file, file.info(file)$size)
    doc <- htmlParse(svg)
    p <- xpathSApply(doc, "//polygon", xmlGetAttr, "points")
    
    if(is_empty(p)){
      #do nothing
    } else {
      # Convert them to numbers
      roiList <- lapply( strsplit(p, " "), function(u) 
        matrix(as.numeric(unlist(strsplit(u, ","))),ncol=2,byrow=TRUE) )
      
      Output <- roiList %>%
        imap_dfr(~ GetXY(.y, .x))
      
      Output$X[Output$X > 1023] <- 1023
      Output$Y[Output$Y > 1023] <- 1023
      Output$X[Output$X < 0] <- 0
      Output$Y[Output$Y < 0] <- 0
      
      for(i in c(1:length(roiList))){
        activeROI <- filter(Output, ShapeNum == i)
        
        #Make string of X and Y coordinate separated by commas. Can only be integer values
        xStr <- paste(activeROI$X, sep="", collapse=",")
        yStr <- paste(activeROI$Y, sep="", collapse=",")
        
        polyWidth <- max(activeROI$X) - min(activeROI$X) + 1
        polyHeight <- max(activeROI$Y) - min(activeROI$Y) + 1
        polyXCoord <- min(activeROI$X)
        polyYCoord <- min(activeROI$Y)
        
        #taskClip
        taskClipToAdd <- taskClip %>%
          filter(TaskID == 1) %>%
          mutate(TaskID = rpCounter,
                 IndexName = 1,
                 IndexID = '1S',
                 Shape = 6,
                 X = polyXCoord,
                 Y = polyYCoord,
                 Width = polyWidth,
                 Height = polyHeight,
                 XCnt = length(activeROI$X),
                 YCnt = length(activeROI$Y),
                 XList = xStr,
                 YList = yStr)
        taskClip <- rbind(taskClip,taskClipToAdd)
        
        #taskInfo
        #repeats enable/disable and set number - 
        #modecheck = 0 even when enabled?
        #Line mode = 2
        taskInfoToAdd <- taskInfo %>%
          filter(TaskID == 1) %>%
          mutate(TaskID = rpCounter,
                 ZPos = zHeight,
                 NumberOfClip = 1,
                 FilterModeLineFrame = if(numRepeats > 1){0} else {2}, 
                 FilterModeNum = numRepeats)
        taskInfo <- rbind(taskInfo,taskInfoToAdd)
        
        #taskBar
        
        #need to figure out minimum time for RPdelay
        betweenRPdelay <- 100 #ms
        frameTime <- frameTimeEstimate(polyWidth,polyHeight) #full frame = 3300 ms
        terminateDuration <- 300 #ms 
        
        mts <- taskBar$TerminateEnd[rpCounter]+betweenRPdelay
        mte <- mts+frameTime
        ps <- taskBar$TerminateEnd[rpCounter]
        pe <- mts
        ts <- mte
        te <- ts+terminateDuration
        
        taskBarToAdd <- taskBar %>%
          filter(TASKID == 1) %>%
          mutate(TASKID = rpCounter,
                 MainTimeStart = mts,
                 MainTimeEnd = mte,
                 PrepareStart = ps,
                 PrepareEnd = pe,
                 TerminateStart = ts,
                 TerminateEnd = te)
        taskBar <- rbind(taskBar,taskBarToAdd)
        
        #channelInfo
        channelInfoToAdd <- channelInfo %>%
          filter(TaskID == 1) %>%
          mutate(TaskID = rpCounter)
        channelInfo <- rbind(channelInfo,channelInfoToAdd)
        
        #taskLaser
        taskLaserToAdd <- taskLaser %>%
          filter(TaskID == 1) %>%
          mutate(TaskID = rpCounter)
        taskLaser <- rbind(taskLaser,taskLaserToAdd)
        
        #taskScaninfo
        taskScaninfoToAdd <- taskScaninfo %>%
          filter(TaskID == 1) %>%
          mutate(TaskID = rpCounter,
                 ClipScanSizeX = if(taskClip$Width[rpCounter] < 1023){taskClip$Width[rpCounter] + 1} else {1024},
                 ClipScanSizeY = if(taskClip$Height[rpCounter] < 1023){taskClip$Height[rpCounter] + 3} else {1024})
        taskScaninfo <- rbind(taskScaninfo,taskScaninfoToAdd)
        
        #taskMatlinfo
        taskMatlinfoToAdd <- taskMatlinfo %>%
          filter(TaskID == 1) %>%
          mutate(TaskID = rpCounter,
                 Num = rpCounter,
                 XIndex = round(4.9697*taskScaninfo$ClipScanSizeX[rpCounter]),
                 YIndex = round(4.9697*taskScaninfo$ClipScanSizeY[rpCounter]))
        taskMatlinfo <- rbind(taskMatlinfo,taskMatlinfoToAdd)
        
        #taskScanrange
        taskScanrangeToAdd <- taskScanrange %>%
          filter(TaskID == 1) %>%
          mutate(TaskID = rpCounter)
        taskScanrange <- rbind(taskScanrange,taskScanrangeToAdd)
        
        
        rpCounter = rpCounter + 1
      }
    }
    zHeight = zHeight + zStepSize
  }
  
  #important, updates end time in db
  updateString <- paste("UPDATE OTHERITEMS_INITVAL SET DataValue = ",
                        taskBar$TerminateEnd[dim(taskBar)[1]],
                        " WHERE PropertyLeafName = 'EndTime'",sep = "")
  sqlQuery( db , updateString)
  
  sqlSave(db, filter(taskClip,TaskID > 1), tablename = "IMAGING_TASK_CLIP", append = TRUE, rownames = FALSE)
  sqlSave(db, filter(taskInfo,TaskID > 1), tablename = "IMAGING_TASK_INFO", append = TRUE, rownames = FALSE)
  sqlSave(db, filter(channelInfo,TaskID > 1), tablename = "IMAGING_CHANNEL_INFO", append = TRUE, rownames = FALSE)
  sqlSave(db, filter(taskLaser,TaskID > 1), tablename = "IMAGING_TASK_LASER", append = TRUE, rownames = FALSE)
  sqlSave(db, filter(taskMatlinfo,TaskID > 1), tablename = "IMAGING_TASK_MATLINFO", append = TRUE, rownames = FALSE)
  sqlSave(db, filter(taskScaninfo,TaskID > 1), tablename = "IMAGING_TASK_SCANINFO", append = TRUE, rownames = FALSE)
  sqlSave(db, filter(taskScanrange,TaskID > 1), tablename = "IMAGING_TASK_SCANRANGE", append = TRUE, rownames = FALSE)
  sqlSave(db, filter(taskBar,TASKID > 1), tablename = "IMAGING_TASK_BAR", append = TRUE, rownames = FALSE)
  odbcCloseAll()
  
  estimatedTime <- if(taskBar$TerminateEnd[dim(taskBar)[1]] > 2000*numPolygons){
    taskBar$TerminateEnd[dim(taskBar)[1]]
  } else{
    2000*numPolygons
  }
  print(paste("Estimated protocol time:", estimatedTime/60/1000, "min"))
}
