#Clean the R environment
rm(list = ls())

#Set working Directory in which the data exist
setwd("D:/Project")
getwd()

#installing all the required packages for model development and preprocessing
install.packages(c("ggplot2","lsr","corrgram","rpart","DataCombine","DMwR","rattle","mltools","pROC","randomForest","inTrees"
                   ,"usdm","Metrics"))
x = c("ggplot2","lsr","corrgram","rpart","DataCombine","DMwR","rattle","mltools","pROC","randomForest","inTrees"
     ,"usdm","Metrics")

#Load XLSX library
#install.packages("xlsx")
library("xlsx")

#loading the data into R environment
df = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1)

#observe the class of df object
class(df)

#Observe the dimension of df
dim(df)

#get the names of variable
colnames(df)

#checking the data types of data
str(df)

#EXPLORATORY DATA ANALYSIS

#note that all the variables of df are of numeric type

#Now we have to observe and explore which are the variables that are continuous and which are categorical variables.

#conversion of required data types into categorical variables a/c to the data
df$ID = as.factor(as.character(df$ID))
df$Day.of.the.week = as.factor(as.character(df$Day.of.the.week))
df$Education = as.factor(as.character(df$Education))
df$Social.drinker = as.factor(as.character(df$Social.drinker))
df$Social.smoker = as.factor(as.character(df$Social.smoker))
df$Reason.for.absence = as.factor(as.character(df$Reason.for.absence))
df$Seasons = as.factor(as.character(df$Seasons))
df$Month.of.absence = as.factor(as.character(df$Month.of.absence))
df$Disciplinary.failure = as.factor(as.character(df$Disciplinary.failure))

#MISSING VALUE ANALYSIS
sum(is.na(df$Transportation.expense))
missing_value = data.frame(apply(df,2,function(x){sum(is.na(x))}))
names(missing_value)[1] = "Missing_data"
missing_value$Variables = row.names(missing_value)
row.names(missing_value) = NULL

#Rearange the columns
missing_value = missing_value[, c(2,1)]

#place the variables according to their number of missing values.
missing_value = missing_value[order(-missing_value$Missing_data),]

#Calculate the missing value percentage
missing_value$percentage = (missing_value$Missing_data/nrow(df) )* 100

#Store the missing value information in a csv file
write.csv(missing_value,"Project_Missing_value.csv", row.names = F)

#Since the calculated missing percentage is less than 5% 
#Thus there is no need to drop any variable due to less number of data
#Now generate a missing value in the dataset and try various imputation methods on it
#make a backup for data object

#df_back = df

## create a missing value in any variable

#*******
#df[79,6] = 361
#ACTUTAL vALUE  = 361
#Median = 225
#KNN = 315
#Mean  = 220


##Now apply mean method to impute this generated value and observe the result
#df$Transportation.expense[is.na(df$Transportation.expense)]  =  mean(df$Transportation.expense, na.rm = T)

#sum(is.na(df))

##Now apply median method to impute the missing value
#df$Transportation.expense[is.na(df$Transportation.expense)] = median(df$Transportation.expense, na.rm = T)

#now apply KNN method to impute
#install.packages('DMwR')
library(DMwR)
df = knnImputation(df,k=3)
sum(is.na(df))

#Both median and KNN are giving same result we can choose any one for furthur implementation
#On analysis of different data using median and KNN i find that KNN is more accurate than median for imputation

#Now load the data again and impute the missing value by KNN
#Check presence of missing values once to confirm

apply(df,2, function(x){sum(is.na(x))})

#NO missing value is found
#create subset of the dataset which have only numeric varaiables

numeric_index = sapply(df, is.numeric)
numeric_data = df[,numeric_index]
numeric_data = as.data.frame(numeric_data)

n_data = colnames(numeric_data)[-12]
str(df)

#draw the boxplot to detect outliers

#install.packages('ggplot2')
library("ggplot2")

for (i in 1:length(n_data)) {
  assign(paste0("gn",i), ggplot(aes_string( y = (n_data[i]), x= "Absenteeism.time.in.hours") , data = subset(df)) + 
           stat_boxplot(geom = "errorbar" , width = 0.5) +
           geom_boxplot(outlier.color = "red", fill = "grey", outlier.shape = 20, outlier.size = 1, notch = FALSE)+
           theme(legend.position = "bottom")+
           labs(y = n_data[i], x= "Absenteeism.time.in.hours")+
           ggtitle(paste("Boxplot" , n_data[i])))
  #print(i)
}

options(warn = -1)

#Now plotting the plots

gridExtra::grid.arrange(gn1, gn2,gn3, ncol=3)
gridExtra::grid.arrange(gn4,gn5,gn6, ncol=3)
gridExtra::grid.arrange(gn7,gn8,gn9, ncol =3)
gridExtra::grid.arrange(gn10,gn11, ncol =2 )

#calculating the outliers found in each variable and printing them.

for (i in n_data) {
  print(i)
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  print(length(val))
  print(val)
}

#Making each outlier as NA for imputation
for (i in n_data) {
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  df[,i][df[,i] %in% val] = NA
}

#Check number of missing values
sum(is.na(df))

#Impute the values using KNN method
df = knnImputation(df, k=3)

#Check again for missing value if present in case
sum(is.na(df))

#Correlation plot
#install.packages("corrgram")
library(corrgram)
corrgram(na.omit(df))
dim(df)
corrgram(df[,n_data],order = F, upper.panel = panel.pie, text.panel = panel.txt, main = "correlation plot" )

#Here we can see in correlation plot that Body Mass Index and weight are highly positively correlated
#I am removing the Body Mass Index variable because weight is a basic variable.

# Select the relevant numerical features
 
num_var = c("Transportation.expense", "Distance.from.Residence.to.work", "Service.time", 
            "Age", "Work.load.Average.day","Hit.target", "Son", "Pet", "Weight", "Height")

#install.packages("lsr")

library("lsr")

anova_test = aov(Absenteeism.time.in.hours ~ ID + Day.of.the.week + Education + 
                   Social.smoker + Social.drinker + Reason.for.absence + Seasons + Month.of.absence + Disciplinary.failure, data = df)
summary(anova_test)

#Dimension reduction
data = subset(df,select = -c(ID,Body.mass.index, Education, Social.smoker, Social.drinker, Seasons, Disciplinary.failure))

#draw histogram on random variables to check if the distributions are normal

hist(data$Hit.target)
hist(data$Work.load.Average.day.)
hist(data$Son)
hist(data$Weight)
hist(data$Transportation.expense)
hist(data$Distance.from.Residence.to.Work)
hist(data$Service.time)
hist(data$Age)
hist(data$Pet)

#THE variables are not seems as normally distributed thus we have to perform Normalisation instead of standardisation
#Now select numerical variables from data_selected Object to perform normalization

print(sapply(data ,is.numeric))

rm(gn1,gn2,gn3,gn4,gn5,gn6,gn7,gn8,gn9,gn10,gn11)


#Feature Scaling
#num_names object contains all the numerical features of selected features
#Normalisation
rm(cnames)
cnames = c("Transportation.expense", "Distance.from.Residence.to.Work","Service.time","Age","Work.load.Average.day.",
           "Hit.target","Son","Pet","Weight","Height")

for(i in cnames){
  print(i)
  data[,i] = (data[,i] - min(data[,i]))/
    (max(data[,i] - min(data[,i])))
}

#Till here we are having our data being normalised and stored in data_selected object
#Write this cleaned data in csv form 
write.csv(data,"processed.csv", row.names = F)

#________________________________________________________________________________
#install.packages("DataCombine")
library(DataCombine)
rmExcept("data")
colnames(data)
#View(data)

#___________ MODEL DEVELOPMENT _______________________

## Divide the data into train and test data using simple random sampling
#install.packages("rpart")
library("rpart")
train_index = sample(1:nrow(data), 0.8* nrow(data))

train = data[train_index,]
test = data[-train_index,]

#Building the Decision Tree Regression Model

regression_model = rpart(Absenteeism.time.in.hours ~. , data = train, method = "anova")

#########  predicting for the test case
predicted_data = predict(regression_model , test[,-14])

##Calculate RMSE to analyze performance of Decision tree regression model
#RMSE is used here as the data is of time series 

library("DMwR")
regr.eval(test[,13], predicted_data, stats = 'rmse')

##__ Thus in Decision tree regression model the error is 9.10 which tells that our model is 90.9% accurate

##________________RANDOM FOREST MODEL __________________________________##
#install.packages("randomForest")
library("randomForest")
RF_model = randomForest(Absenteeism.time.in.hours~. , train,  ntree=100)

#Extract the rules generated as a result of random Forest model
#install.packages("inTrees")
library("inTrees")
rules_list = RF2List(RF_model)

#Extract rules from rules_list
rules = extractRules(rules_list, train[,-14])
rules[1:2,]

#Convert the rules in readable format
read_rules = presentRules(rules,colnames(train))
read_rules[1:2,]

#Determining the rule metric
rule_metric = getRuleMetric(rules, train[,-14], train$Absenteeism.time.in.hours)
rule_metric[1:2,]

#Prediction of the target variable data using the random Forest model
RF_prediction = predict(RF_model,test[,-14])
regr.eval(test[,13], RF_prediction, stats = 'rmse')

#Thus the error rate in Random Forest Model is 8.78% and the accuracy of the model is 100-8.78 = 91.22%

###________________ LINEAR REGRESSION _______________________________
install.packages("usdm")
library("usdm")
LR_data_select = subset(data,select = -c(Reason.for.absence,Day.of.the.week))
colnames(LR_data_select)
vif(LR_data_select[,-11])
vifcor(LR_data_select[,-11], th=0.9)

####Execute the linear regression model over the data
lr_model = lm(Absenteeism.time.in.hours~. , data = train)

summary(lr_model)
colnames(test)

#Predict the data 
LR_predict_data = predict(lr_model, test[,1:13])

regr.eval(test[,13], LR_predict_data, stats = 'rmse')


##________ Till here we have implemented Decision Tree, Random Forest and Linear Regression. Among all of these,
##________ Random Forest is having highest accuracy.


#________ Now we have to predict loss for each month ____________

#__ To calculate loss month wise we need to include month of absence variable again in our data set 
# LOSS = Work.load.average.per.day * Absenteeism.time.in.hours

colnames(data)
data$loss = data$Work.load.Average.day. * data$Absenteeism.time.in.hours

i=1

#______________________________________________________________________________

# NOW calculate Month wise loss encountered due to absenteeism of employees 

#Calculate loss in january
loss_jan = as.data.frame(data$loss[data$Month.of.absence %in% 1])
names(loss_jan)[1] = "Loss"
sum(loss_jan[1])
write.csv(loss_jan,"jan_loss.csv", row.names = F)

#Calculate loss in febreuary
loss_feb = as.data.frame(data$loss[data$Month.of.absence %in% 2])
names(loss_feb)[1] = "Loss"
sum(loss_feb[1])
write.csv(loss_feb,"feb_loss.csv", row.names = F)

#Calculate loss in march

loss_march = as.data.frame(data$loss[data$Month.of.absence %in% 3])
names(loss_march)[1] = "Loss"
sum(loss_march[1])
write.csv(loss_march,"march_loss.csv", row.names = F)

#Calculate loss in april

loss_apr = as.data.frame(data$loss[data$Month.of.absence %in% 4])
names(loss_apr)[1] = "Loss"
sum(loss_apr[1])
write.csv(loss_apr,"apr_loss.csv", row.names = F)


#calculate loss in may

loss_may = as.data.frame(data$loss[data$Month.of.absence %in% 5])
names(loss_may)[1] = "Loss"
sum(loss_may[1])
write.csv(loss_may,"may_loss.csv", row.names = F)


#calculate in june

loss_jun = as.data.frame(data$loss[data$Month.of.absence %in% 6])
names(loss_jun)[1] = "Loss"
sum(loss_jun[1])
write.csv(loss_jun,"jun_loss.csv", row.names = F)

#Calculate loss in july

loss_jul = as.data.frame(data$loss[data$Month.of.absence %in% 7])
names(loss_jul)[1] = "Loss"
sum(loss_jul[1])
write.csv(loss_jul,"jul_loss.csv", row.names = F)

#calculate loss in august

loss_aug = as.data.frame(data$loss[data$Month.of.absence %in% 8])
names(loss_aug)[1] = "Loss"
sum(loss_aug[1])
write.csv(loss_aug,"aug_loss.csv", row.names = F)

#Calculate loss in september

loss_sep = as.data.frame(data$loss[data$Month.of.absence %in% 9])
names(loss_sep)[1] = "Loss"
sum(loss_sep[1])
write.csv(loss_sep,"sep_loss.csv", row.names = F)

#calculate loss in october

loss_oct = as.data.frame(data$loss[data$Month.of.absence %in% 10])
names(loss_oct)[1] = "Loss"
sum(loss_oct[1])
write.csv(loss_oct,"oct_loss.csv", row.names = F)

#calculate loss in november

loss_nov = as.data.frame(data$loss[data$Month.of.absence %in% 2])
names(loss_nov)[1] = "Loss"
sum(loss_nov[1])
write.csv(loss_nov,"nov_loss.csv", row.names = F)

#calculate loss in december
loss_dec = as.data.frame(data$loss[data$Month.of.absence %in% 2])
names(loss_dec)[1] = "Loss"
sum(loss_dec[1])
write.csv(loss_dec,"dec_loss.csv", row.names = F)

