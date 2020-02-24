#!/bin/sh -x
#Above is shebang to run script in Shell with verbose
#Shell script to execute ETL pipeline

#Takes exactly two required arguments, Hadoop User DB name and email address for status update
#Below block makes sure exactly two arguments are provided for process to continue
if [[ $# -lt 2 ]]
then
  echo "READ ME: All required arguments not provided, two needed. Check doc for ref."
  exit
elif [[ $# -gt 2 ]]
then
  echo "READ ME: More than two arguments provided. Check doc for ref"
  exit
fi

#Delete all quotation marks except for in field quotations in weekly upload flat file
#Example: name = Robert "Bob" (keep these quotations)
sed -i -r 's/\t\"/\t/g; s/\"\t/\t/g; s/\"\"/\"/g; s/^\"//g; s/\"$//g;' [UNIX weekly flat file path]

#Catch error block checks whether above command executed successfully
if [[ $? -gt 0 ]]
then
  echo "READ ME: 'sed' command didn't execute successfully. Make sure file path is accurate and exists."
  exit
fi

#Hadoop commands to move weekly flat file from UNIX to Hadoop environment

#Drop existing HDFS directory and its data
hdfs dfs -rm -R [hadoop db path]
#Catch error block checks whether above command executed successfully
if [[ $? -gt 0 ]]
then 
  echo "READ ME: 'rm' command didn't execute successfully. Make sure Hadoop DB name/path is accurate."
  exit
fi

#Create HDFS directory to put data from weekly upload flat file
hdfs dfs -mkdir [hadoop db path]
#Catch error block checks whether above command executed successfully
if [[ $? -gt 0 ]]
then 
  echo "READ ME: 'mkdir' command didn't execute successfully. Make sure Hadoop DB name/path is accurate."
  exit
fi

#Put data from weekly upload flat file in HDFS
hdfs dfs -put [UNIX weekly flat file path] [hadoop db path]
#Catch error block checks whether above command executed successfully
if [[ $? -gt 0 ]]
then
  echo "READ ME: 'put' command didn't execute successfully. Make sure UNIX and Hadoop DB path is accurate."
  exit
fi

#Concate multiple SQL queries into one seqential ETL pipeline for efficient run  
cat [SQL files path]*.sql > [concat file path].sql
#Catch error block checks whether above command executed successfully
if [[ $? -gt 0 ]]
then
  echo "READ ME: 'cat' command didn't execute successfully. Further investigation is required."
  exit
fi

#Run ETL on Hive using below command
#Setting beeline session variables for databases
#Queries will be calling these variables during runtime
beeline --hivevar trgt_db=target_db_name       \
        --hivevar source_1=source_1_name \
        --hivevar source_2=source_2_name  \
        --hivevar hadoop_db=$1 \
         -f [concat file path].sql

#Sends an email notification when update is finished
echo "Refresh process complete. Examine the log file to confirm the successfull run." | mailx -r "$2" -s "Hadoop Refresh Complete" "$2"
#Catch error block checks whether above command executed successfully
if [[ $? -gt 0 ]]
then
  echo "READ ME: 'mailx' command didn't execute successfully. Make sure email address is accurate."
  exit
fi
