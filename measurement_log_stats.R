rm(list = ls())
library(openxlsx)

# change this to the path of the file you want to analyze (can be any path)
f1 = "/Users/timjun/Desktop/USVI_2026/combined_data/0607_R20/all_R20/combined_R20_06072026_MF_AM_noGPS.xlsx"

# don't change anything after this point
df1 = read.xlsx(f1)


# calculate percent for GPS and temp by summing number of seconds from start to end

df1$time2 = as.POSIXct(paste(df1$date, df1$time, sep=" "),
                        format = "%m/%d/%Y %H:%M:%S", tz = "UTC")

# GPS and temperature obs are valid if not NA and not 0
df1$gps_valid = ifelse(
  !is.na(df1$latitude) & !is.na(df1$longitude) & df1$latitude != 0 & df1$longitude != 0,
  "Y", "N"
)
df1$temp_valid = ifelse(
  !is.na(df1$temp_probe) & df1$temp_probe != 0 ,
  "Y", "N"
)

# create format to hold summary stats
out=data.frame(matrix(ncol=14,nrow=1))
colnames(out) = c("Start","End","GPS_pct","Temp_pct","Tmax","Tmin","Dew_point_max",
                  "Dew_point_min","RH_max","RH_min","HI_max","HI_min","Ambient_max",
                  "Ambient_min")

# populate summary stats
out$Start = format(min(df1$time2), "%H:%M:%S")
out$End = format(max(df1$time2), "%H:%M:%S")  

duration = as.numeric(difftime(max(df1$time2),min(df1$time2),units="secs"))
valid_GPS_count = sum(df1$gps_valid=="Y")
out$GPS_pct = valid_GPS_count/duration

valid_temp_count = sum(df1$temp_valid=="Y")
out$Temp_pct = valid_temp_count/duration

out$Tmax = max(df1$temp_probe)
out$Tmin = min(df1$temp_probe)

out$Dew_point_max = max(df1$dew_point)
out$Dew_point_min = min(df1$dew_point)

out$RH_max = max(df1$humidity)
out$RH_min = min(df1$humidity)

out$HI_max = max(df1$heat_index)
out$HI_min = min(df1$heat_index)

out$Ambient_max = max(df1$ambient_light)
out$Ambient_min = min(df1$ambient_light)

# can either run code below to write summary stats to file
# or can simply copy and paste dateframe out to MeasurementLog
# output name is built from the input file: stats_<input file name>.csv
out_name = paste("stats_", tools::file_path_sans_ext(basename(f1)), ".csv", sep="")
write.csv(out,out_name)

# show the result and where it was saved so you can see it worked
# t() flips it so each stat is on its own line (reads better in a narrow terminal)
print(t(out))
cat("Saved:", normalizePath(out_name), "\n")

