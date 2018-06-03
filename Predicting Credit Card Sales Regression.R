setwd("G:\\Current\\SAS\\SASUniversityEdition\\myfolders\\MEGA Case\\Linear Regression in R")
options(java.parameters = "- Xmx1024m")

library(XLConnect)
Linear <- readWorksheetFromFile("Linear Regression Case.xlsx",sheet=1)

# Identifying Outliers
mystats <- function(x) {
  nmiss<-sum(is.na(x))
  a <- x[!is.na(x)]
  m <- mean(a)
  n <- length(a)
  s <- sd(a)
  min <- min(a)
  p1<-quantile(a,0.01)
  p5<-quantile(a,0.05)
  p10<-quantile(a,0.10)
  q1<-quantile(a,0.25)
  q2<-quantile(a,0.5)
  q3<-quantile(a,0.75)
  p90<-quantile(a,0.90)
  p95<-quantile(a,0.95)
  p99<-quantile(a,0.99)
  max <- max(a)
  UC <- m+2*s
  LC <- m-2*s
  outlier_flag<- max>UC | min<LC
  return(c(n=n, nmiss=nmiss, outlier_flag=outlier_flag, mean=m, stdev=s,min = min, p1=p1,p5=p5,p10=p10,q1=q1,q2=q2,q3=q3,p90=p90,p95=p95,p99=p99,max=max, UC=UC, LC=LC ))
}


Num_Vars <- c("age",
              "card2items",
              "carditems",
              "cardmon",
              "cardten",
              "carvalue",
              "commutetime",
              "creddebt",
              "debtinc",
              "ed",
              "equipmon",
              "equipten",
              "hourstv",
              "income",
              "lncardmon",
              "lncardten",
              "lncreddebt",
              "lnequipmon",
              "lnequipten",
              "lninc",
              "lnlongmon",
              "lnlongten",
              "lnothdebt",
              "lntollmon",
              "lntollten",
              "lnwiremon",
              "lnwireten",
              "longmon",
              "longten",
              "othdebt",
              "pets",
              "pets_birds",
              "pets_cats",
              "pets_dogs",
              "pets_freshfish",
              "pets_reptiles",
              "pets_saltfish",
              "pets_small",
              "reside",
              "spoused",
              "tenure",
              "tollmon",
              "tollten",
              "wiremon",
              "wireten")


Cat_Vars <- c("cars",
              "carown",
              "cartype",
              "carcatvalue",
              "carbought",
              "reason",
              "polview",
              "card",
              "cardtype",
              "cardbenefit",
              "card2",
              "card2type",
              "card2benefit",
              "bfast",
              "internet",
              "empcat",
              "inccat",
              "jobsat",
              "spousedcat",
              "hometype",
              "addresscat",
              "commute",
              "commutecat",
              "cardtenurecat",
              "card2tenurecat",
              "jobcat",
              "townsize",
              "reason",
              "region")
# Outlier              
Outliers<-t(data.frame(apply(Linear[Num_Vars], 2, mystats)))
View(Outliers)
Step1 <- Linear
Step1$age[Step1$age>79]<-79
Step1$card2items[Step1$card2items>11]<-11
Step1$carditems[Step1$carditems>19]<-19
Step1$cardmon[Step1$cardmon>64.25]<-64.25
Step1$cardten[Step1$cardten>4011.2]<-4011.2
Step1$carvalue[Step1$carvalue>92.001]<-92.001
Step1$commutetime[Step1$commutetime>40.03]<-40.03
Step1$creddebt[Step1$creddebt>14.280358]<-14.280358
Step1$debtinc[Step1$debtinc>29.2]<-29.2
Step1$ed[Step1$ed>21]<-21
Step1$equipmon[Step1$equipmon>63.3005]<-63.3005
Step1$equipten[Step1$equipten>3679.4575]<-3679.4575
Step1$hourstv[Step1$hourstv>31]<-31
Step1$income[Step1$income>272.01]<-272.01
Step1$lncardmon[Step1$lncardmon>4.239162]<-4.239162
Step1$lncardten[Step1$lncardten>8.392151]<-8.392151
Step1$lncreddebt[Step1$lncreddebt>2.65891]<-2.65891
Step1$lnequipmon[Step1$lnequipmon>4.269466]<-4.269466
Step1$lnequipten[Step1$lnequipten>8.369037]<-8.369037
Step1$lninc[Step1$lninc>5.605839]<-5.605839
Step1$lnlongmon[Step1$lnlongmon>4.177475]<-4.177475
Step1$lnlongten[Step1$lnlongten>8.452988]<-8.452988
Step1$lnothdebt[Step1$lnothdebt>3.180802]<-3.180802
Step1$lntollmon[Step1$lntollmon>4.190524]<-4.190524
Step1$lntollten[Step1$lntollten>8.429812]<-8.429812
Step1$lnwiremon[Step1$lnwiremon>4.577186]<-4.577186
Step1$lnwireten[Step1$lnwireten>8.690117]<-8.690117
Step1$longmon[Step1$longmon>65.201]<-65.201
Step1$longten[Step1$longten>4689.066]<-4689.066
Step1$othdebt[Step1$othdebt>24.06426]<-24.06426
Step1$pets[Step1$pets>13]<-13
Step1$pets_birds[Step1$pets_birds>3]<-3
Step1$pets_cats[Step1$pets_cats>3]<-3
Step1$pets_dogs[Step1$pets_dogs>3]<-3
Step1$pets_freshfish[Step1$pets_freshfish>11]<-11
Step1$pets_reptiles[Step1$pets_reptiles>2]<-2
Step1$pets_saltfish[Step1$pets_saltfish>2]<-2
Step1$pets_small[Step1$pets_small>3]<-3
Step1$reside[Step1$reside>6]<-6
Step1$spoused[Step1$spoused>20]<-20
Step1$tenure[Step1$tenure>72]<-72
Step1$tollmon[Step1$tollmon>58.7525]<-58.7525
Step1$tollten[Step1$tollten>3977.2705]<-3977.2705
Step1$wiremon[Step1$wiremon>78.304]<-78.304
Step1$wireten[Step1$wireten>4530.186]<-4530.186

#Missing Value Imputation
Missing<-t(data.frame(apply(Step1[Num_Vars], 2, mystats)))
View(Missing)
Step1$age[which(is.na(Step1$age))]<-47.0256
Step1$card2items[which(is.na(Step1$card2items))]<-4.6564
Step1$carditems[which(is.na(Step1$carditems))]<-10.1688
Step1$cardmon[which(is.na(Step1$cardmon))]<-15.26695
Step1$cardten[which(is.na(Step1$cardten))]<-707.2390956
Step1$carvalue[which(is.na(Step1$carvalue))]<-23.20223
Step1$commutetime[which(is.na(Step1$commutetime))]<-25.3160264
Step1$creddebt[which(is.na(Step1$creddebt))]<-1.7582311
Step1$debtinc[which(is.na(Step1$debtinc))]<-9.91152
Step1$ed[which(is.na(Step1$ed))]<-14.5348
Step1$equipmon[which(is.na(Step1$equipmon))]<-12.908715
Step1$equipten[which(is.na(Step1$equipten))]<-463.398395
Step1$hourstv[which(is.na(Step1$hourstv))]<-19.6266
Step1$income[which(is.na(Step1$income))]<-53.6299
Step1$lncardmon[which(is.na(Step1$lncardmon))]<-2.9075331
Step1$lncardten[which(is.na(Step1$lncardten))]<-6.4239255
Step1$lncreddebt[which(is.na(Step1$lncreddebt))]<--0.1345418
Step1$lnequipmon[which(is.na(Step1$lnequipmon))]<-3.5990319
Step1$lnequipten[which(is.na(Step1$lnequipten))]<-6.7456924
Step1$lninc[which(is.na(Step1$lninc))]<-3.6970492
Step1$lnlongmon[which(is.na(Step1$lnlongmon))]<-2.2863219
Step1$lnlongten[which(is.na(Step1$lnlongten))]<-5.6088459
Step1$lnothdebt[which(is.na(Step1$lnothdebt))]<-0.6933673
Step1$lntollmon[which(is.na(Step1$lntollmon))]<-3.2418264
Step1$lntollten[which(is.na(Step1$lntollten))]<-6.5833077
Step1$lnwiremon[which(is.na(Step1$lnwiremon))]<-3.6031934
Step1$lnwireten[which(is.na(Step1$lnwireten))]<-6.8059049
Step1$longmon[which(is.na(Step1$longmon))]<-13.26915
Step1$longten[which(is.na(Step1$longten))]<-694.3159696
Step1$othdebt[which(is.na(Step1$othdebt))]<-3.5220976
Step1$pets[which(is.na(Step1$pets))]<-3.0492
Step1$pets_birds[which(is.na(Step1$pets_birds))]<-0.106
Step1$pets_cats[which(is.na(Step1$pets_cats))]<-0.4904
Step1$pets_dogs[which(is.na(Step1$pets_dogs))]<-0.3828
Step1$pets_freshfish[which(is.na(Step1$pets_freshfish))]<-1.8348
Step1$pets_reptiles[which(is.na(Step1$pets_reptiles))]<-0.05
Step1$pets_saltfish[which(is.na(Step1$pets_saltfish))]<-0.0226
Step1$pets_small[which(is.na(Step1$pets_small))]<-0.1028
Step1$reside[which(is.na(Step1$reside))]<-2.1942
Step1$spoused[which(is.na(Step1$spoused))]<-6.0954
Step1$tenure[which(is.na(Step1$tenure))]<-38.2048
Step1$tollmon[which(is.na(Step1$tollmon))]<-13.140075
Step1$tollten[which(is.na(Step1$tollten))]<-570.130195
Step1$wiremon[which(is.na(Step1$wiremon))]<-10.53027
Step1$wireten[which(is.na(Step1$wireten))]<-409.96002

Mode <- function(x) {
  ux <- unique(x)
  ux[which.max(tabulate(match(x, ux)))]
}

categoricals_mode<-t(data.frame(apply(Step1[Cat_Vars], 2, Mode)))
View(categoricals_mode)

Step1$cars[which(is.na(Step1$cars))]<-2
Step1$carown[which(is.na(Step1$carown))]<-1
Step1$cartype[which(is.na(Step1$cartype))]<-0
Step1$carcatvalue[which(is.na(Step1$carcatvalue))]<-1
Step1$carbought[which(is.na(Step1$carbought))]<-0
Step1$reason[which(is.na(Step1$reason))]<-9
Step1$polview[which(is.na(Step1$polview))]<-4
Step1$card[which(is.na(Step1$card))]<-4
Step1$cardtype[which(is.na(Step1$cardtype))]<-4
Step1$cardbenefit[which(is.na(Step1$cardbenefit))]<-3
Step1$card2[which(is.na(Step1$card2))]<-3
Step1$card2type[which(is.na(Step1$card2type))]<-4
Step1$card2benefit[which(is.na(Step1$card2benefit))]<-4
Step1$bfast[which(is.na(Step1$bfast))]<-3
Step1$internet[which(is.na(Step1$internet))]<-0
Step1$empcat[which(is.na(Step1$empcat))]<-2
Step1$inccat[which(is.na(Step1$inccat))]<-2
Step1$jobsat[which(is.na(Step1$jobsat))]<-3
Step1$spousedcat[which(is.na(Step1$spousedcat))]<--1
Step1$hometype[which(is.na(Step1$hometype))]<-1
Step1$addresscat[which(is.na(Step1$addresscat))]<-3
Step1$commute[which(is.na(Step1$commute))]<-1
Step1$commutecat[which(is.na(Step1$commutecat))]<-1
Step1$cardtenurecat[which(is.na(Step1$cardtenurecat))]<-5
Step1$card2tenurecat[which(is.na(Step1$card2tenurecat))]<-5
Step1$jobcat[which(is.na(Step1$jobcat))]<-2
Step1$townsize[which(is.na(Step1$townsize))]<-1
Step1$region[which(is.na(Step1$region))]<-5

#Creating Dummy Variables
Step2 <- Step1

Step2$cars<- factor(Step2$cars)
Step2$cars<- factor(Step2$cars)
Step2$carown<- factor(Step2$carown)
Step2$cartype<- factor(Step2$cartype)
Step2$carcatvalue<- factor(Step2$carcatvalue)
Step2$carbought<- factor(Step2$carbought)
Step2$reason<- factor(Step2$reason)
Step2$polview<- factor(Step2$polview)
Step2$card<- factor(Step2$card)
Step2$cardtype<- factor(Step2$cardtype)
Step2$cardbenefit<- factor(Step2$cardbenefit)
Step2$card2<- factor(Step2$card2)
Step2$card2type<- factor(Step2$card2type)
Step2$card2benefit<- factor(Step2$card2benefit)
Step2$bfast<- factor(Step2$bfast)
Step2$internet<- factor(Step2$internet)
Step2$empcat<- factor(Step2$empcat)
Step2$inccat<- factor(Step2$inccat)
Step2$jobsat<- factor(Step2$jobsat)
Step2$spousedcat<- factor(Step2$spousedcat)
Step2$hometype<- factor(Step2$hometype)
Step2$addresscat<- factor(Step2$addresscat)
Step2$commute<- factor(Step2$commute)
Step2$commutecat<- factor(Step2$commutecat)
Step2$cardtenurecat<- factor(Step2$cardtenurecat)
Step2$card2tenurecat<- factor(Step2$card2tenurecat)
Step2$jobcat<- factor(Step2$jobcat)
Step2$townsize<- factor(Step2$townsize)
Step2$reason<- factor(Step2$reason)
Step2$region<- factor(Step2$region)


#Checking the need of log
Step2$Y <- Step2$cardspent+Step2$card2spent
hist(Step2$Y)
hist(log(Step2$Y))
Step2$ln_Y<-log(Step2$Y)

# Variable Reduction (Factor Analysis)
Step_nums <- Step2[Num_Vars]
corrm<- cor(Step_nums)    
eigen(corrm)$values

require(dplyr)

eigen_values <- mutate(data.frame(eigen(corrm)$values)
                       ,cum_sum_eigen=cumsum(eigen.corrm..values)
                       , pct_var=eigen.corrm..values/sum(eigen.corrm..values)
                       , cum_pct_var=cum_sum_eigen/sum(eigen.corrm..values))

write.csv(eigen_values, "EigenValues.csv")

require(psych)
FA<-fa(r=corrm, 16, rotate="varimax", fm="ml")  
#SORTING THE LOADINGS
FA_SORT<-fa.sort(FA)                                         
FA_SORT$loadings


Loadings<-data.frame(FA_SORT$loadings[1:ncol(Step_nums),])
write.csv(Loadings, "loadings.csv")

# Variable Reduction (Binary Categorical Variables-Indipendent 2 sample T test)
Step1$Y <- Step2$cardspent+Step2$card2spent
Step1$ln_Y<-log(Step1$Y)
t.test(Step1$ln_Y~Step1$union, data=Step1)
t.test(Step1$ln_Y~Step1$retire, data=Step1)
t.test(Step1$ln_Y~Step1$homeown, data=Step1)
t.test(Step1$ln_Y~Step1$carbuy, data=Step1)
t.test(Step1$ln_Y~Step1$commutecar, data=Step1)
t.test(Step1$ln_Y~Step1$commutemotorcycle, data=Step1)
t.test(Step1$ln_Y~Step1$commuterail, data=Step1)
t.test(Step1$ln_Y~Step1$commutepublic, data=Step1)
t.test(Step1$ln_Y~Step1$commutebike, data=Step1)
t.test(Step1$ln_Y~Step1$commutewalk, data=Step1)
t.test(Step1$ln_Y~Step1$commutenonmotor, data=Step1)
t.test(Step1$ln_Y~Step1$telecommute, data=Step1)
t.test(Step1$ln_Y~Step1$polparty, data=Step1)
t.test(Step1$ln_Y~Step1$polcontrib, data=Step1)
t.test(Step1$ln_Y~Step1$vote, data=Step1)
t.test(Step1$ln_Y~Step1$cardfee, data=Step1)
t.test(Step1$ln_Y~Step1$active, data=Step1)
t.test(Step1$ln_Y~Step1$churn, data=Step1)
t.test(Step1$ln_Y~Step1$tollfree, data=Step1)
t.test(Step1$ln_Y~Step1$equip, data=Step1)
t.test(Step1$ln_Y~Step1$callcard, data=Step1)
t.test(Step1$ln_Y~Step1$wireless, data=Step1)
t.test(Step1$ln_Y~Step1$multline, data=Step1)
t.test(Step1$ln_Y~Step1$voice, data=Step1)
t.test(Step1$ln_Y~Step1$pager, data=Step1)
t.test(Step1$ln_Y~Step1$callid, data=Step1)
t.test(Step1$ln_Y~Step1$callwait, data=Step1)
t.test(Step1$ln_Y~Step1$forward, data=Step1)
t.test(Step1$ln_Y~Step1$confer, data=Step1)
t.test(Step1$ln_Y~Step1$ebill, data=Step1)
t.test(Step1$ln_Y~Step1$owntv, data=Step1)
t.test(Step1$ln_Y~Step1$ownvcr, data=Step1)
t.test(Step1$ln_Y~Step1$owndvd, data=Step1)
t.test(Step1$ln_Y~Step1$owncd, data=Step1)
t.test(Step1$ln_Y~Step1$ownpda, data=Step1)
t.test(Step1$ln_Y~Step1$ownpc, data=Step1)
t.test(Step1$ln_Y~Step1$ownipod, data=Step1)
t.test(Step1$ln_Y~Step1$owngame, data=Step1)
t.test(Step1$ln_Y~Step1$ownfax, data=Step1)
t.test(Step1$ln_Y~Step1$news, data=Step1)
t.test(Step1$ln_Y~Step1$response_01, data=Step1)
t.test(Step1$ln_Y~Step1$response_02, data=Step1)
t.test(Step1$ln_Y~Step1$response_03, data=Step1)
t.test(Step1$ln_Y~Step1$default, data=Step1)
t.test(Step1$ln_Y~Step1$marital, data=Step1)
t.test(Step1$ln_Y~Step1$commutecarpool, data=Step1)
t.test(Step1$ln_Y~Step1$commutebus, data=Step1)
t.test(Step1$ln_Y~Step1$card2fee, data=Step1)

# Variable Reduction (Multiple Categorical Variables-Anova)
a1<-aov(ln_Y~cars, data=Step1)	; summary (a1)
a2<-aov(ln_Y~carown, data=Step1)	; 	summary (a2)
a3<-aov(ln_Y~cartype, data=Step1)	; 	summary (a3)
a4<-aov(ln_Y~carcatvalue, data=Step1)	; 	summary (a4)
a5<-aov(ln_Y~carbought, data=Step1)	; 	summary (a5)
a6<-aov(ln_Y~reason, data=Step1)	; 	summary (a6)
a7<-aov(ln_Y~polview, data=Step1)	; 	summary (a7)
a8<-aov(ln_Y~card, data=Step1)	; 	summary (a8)
a9<-aov(ln_Y~cardtype, data=Step1)	; 	summary (a9)
a10<-aov(ln_Y~cardbenefit, data=Step1)	; 	summary (a10)
a11<-aov(ln_Y~card2, data=Step1)	; 	summary (a11)
a12<-aov(ln_Y~card2type, data=Step1)	; 	summary (a12)
a13<-aov(ln_Y~card2benefit, data=Step1)	; 	summary (a13)
a14<-aov(ln_Y~bfast, data=Step1)	; 	summary (a14)
a15<-aov(ln_Y~internet, data=Step1)	; 	summary (a15)
a16<-aov(ln_Y~empcat, data=Step1)	; 	summary (a16)
a17<-aov(ln_Y~inccat, data=Step1)	; 	summary (a17)
a18<-aov(ln_Y~jobsat, data=Step1)	; 	summary (a18)
a19<-aov(ln_Y~spousedcat, data=Step1)	; 	summary (a19)
a20<-aov(ln_Y~hometype, data=Step1)	; 	summary (a20)
a21<-aov(ln_Y~addresscat, data=Step1)	; 	summary (a21)
a22<-aov(ln_Y~commute, data=Step1)	; 	summary (a22)
a23<-aov(ln_Y~commutecat, data=Step1)	; 	summary (a23)
a24<-aov(ln_Y~cardtenurecat, data=Step1)	; 	summary (a24)
a25<-aov(ln_Y~card2tenurecat, data=Step1)	; 	summary (a25)
a26<-aov(ln_Y~jobcat, data=Step1)	; 	summary (a26)
a27<-aov(ln_Y~townsize, data=Step1)	; 	summary (a27)
a28<-aov(ln_Y~reason, data=Step1)	; 	summary (a28)
a29<-aov(ln_Y~region, data=Step1)	; 	summary (a29)

#Splitting data into Training, Validaton and Testing Dataset
set.seed(123)
splitter <- sample(1:nrow(Step2), size = floor(0.70 * nrow(Step2)))
training<-Step2[splitter,]
testing<-Step2[-splitter,]

# Regression
Train_Out <- lm(ln_Y ~ 
                  card+
                  card2+
                  reason+
                  age+
                  lninc+
                  card2items+
                  carditems, 
           data=training)

summary(Train_Out)

require(MASS)
Train_AIC<- stepAIC(Train_Out,direction="both")

library(car)
vif(Train_Out)

#Validation (Train)
t1<-cbind(training, pred_sales = exp(predict(Train_Out)))

#Validation (Test)
t2<-cbind(testing, pred_sales=exp(predict(Train_Out,testing)))


#Decile Analysis Reports - t1(training)

decLocations <- quantile(t1$pred_sales, probs = seq(0.1,0.9,by=0.1))

t1$decile <- findInterval(t1$pred_sales,c(-Inf,decLocations, Inf))

require(sqldf)
t1_DA <- sqldf("select decile, count(decile) as count, avg(pred_sales) as avg_pre_sales,   
               avg(Y) as avg_Actual_sales
               from t1
               group by decile
               order by decile desc")

View(t1_DA)
write.csv(t1_DA,"mydata1_DA.csv")


################################################

# find the decile locations 
decLocations <- quantile(t2$pred_sales, probs = seq(0.1,0.9,by=0.1))

# use findInterval with -Inf and Inf as upper and lower bounds
t2$decile <- findInterval(t2$pred_sales,c(-Inf,decLocations, Inf))

require(sqldf)
t2_DA <- sqldf("select decile, count(decile) as count, avg(pred_sales) as avg_pre_sales,   
               avg(Y) as avg_Actual_sales
               from t2
               group by decile
               order by decile desc")

View(t2_DA)
write.csv(t2_DA,"t2_DA.csv")


