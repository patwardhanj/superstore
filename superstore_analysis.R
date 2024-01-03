setwd("C:\\Users\\amrut\\Downloads\\TheLearnedSage\\Datasets\\Superstore_data")

library(readxl)  #for reading xlsx files
library(writexl) #for writing xlsx files
library(dplyr) #for data cleaning and wrangling
library(psych) #for descriptive statistcs

ss =read_xlsx("superstore_analysis.xlsx") #51389 observations

#---------------PART 1: understand the data--------------------
#how many columns exist?
#what is the data in each of the columns?
#what is the structure of the columns?
#--------------------------------------------------------------
View(ss)
str(ss)  #Ship.Date in character form instead of date

#---------------PART 2: examine the data--------------------
#min-max of each numeric / float / date column
#unique values in each character column
#duplicate rows
#repeating occurrences of columns like customer id, product id etc. to understand 
#process
#--------------------------------------------------------------

#get description to understand min-max
describe(ss)

#observations:
#Shu-sh has no variation in values: looks like all values are 1 (min=1, max=1, range=0)
#There are negative quantities, which don't make sense (min=-11)
#minimum value for sales is 0: which does needs investigation
#there are missing values in 2 columns: Country, Order.Date

#Next: understand unique values in each of the string columns
#this helps in identifying unwanted entries (e.g., 'NAs', "Not Available" etc.)

unique(ss$Category)
unique(ss$Country)
unique(ss$Market)
unique(ss$Order.Priority)
unique(ss$Region)
unique(ss$Ship.Mode)
unique(ss$Sub.Category)
unique(ss$Market2)

#observations: under category: Office Supplies and Office Supplies & Stationary are separate
#are they really separate?
#Order priority has "NA"s

#---------------PART 3: Start cleaning up the data--------------------
#change column data types as needed
#remove / address NAs
#drop inconsistent entries (e.g., -ve prices)
#drop columns not useful for analysis
#investigate uncertain inconsistencies and address them
#--------------------------------------------------------------

#change shipping date & order date to date format
ss$Ship.Date = as.Date(ss$Ship.Date)
ss$Order.Date = as.Date(ss$Order.Date)
str(ss) # both date columns are in order

#drop unwanted column Shu-sh
ss = ss %>% select(-`Shu-Sh`)

#drop negative quantity entries
ss = ss %>% filter(Quantity > 0)

#drop duplicates
ss = ss %>% distinct()

#examine 0 sales value:
x = ss %>% filter(Sales <= 0) #only 1 observation: drop it
ss = ss %>% filter(Sales > 0)

#handle NAs in Order Date
x = ss %>% filter(is.na(Order.Date))
View(x) #looks like there is nothing one can do: drop these 765 rows

ss = ss %>% filter(!(is.na(Order.Date)))

rm(x)
#Handle "NA" as a string from 'Order Priority' column 
#Assign 1 = 'Critical", 2 = 'High', 3 = 'Medium', 4 = 'Low'
#Take average, round and equate back to fill replace 'NA's

op = ss %>% select(Order.Priority) %>% filter(Order.Priority != "NA") %>%
     mutate(priority = case_when(Order.Priority == "Critical" ~ 1,
                                 Order.Priority == "High" ~ 2,
                                 Order.Priority == "Medium" ~ 3,
                                 Order.Priority == "Low" ~ 4))
round(mean(op$priority), digits = 0)
rm(op)
#3 = 'Medium', replace 'NA' by 'Medium' 
ss = ss %>% mutate(Order.Priority = 
                     ifelse(Order.Priority == "NA", "Medium", 
                            Order.Priority)) 
table(ss$Order.Priority)

#check NAs in country columns
x = ss %>% filter(is.na(Country))
View(x) #looks like we can map missing countries based on cities
rm(x)

#separate city and country columns: remove NAs: remove duplicates: merge based on pairs
city_country = ss %>% select(City, Country) %>% 
               filter(!is.na(Country)) %>% distinct() #3693 unique combinations

#check if a city appears multiple times for different countries: as it would 
#cause multiple mappings and one can't be sure
city_country = city_country %>% group_by(City) %>% mutate(ct = n()) %>%
               filter(ct == 1) %>% select (-ct)

#make sure the base sheet (city_country) is the first argument
ss = merge(city_country, ss, by = "City", all.x = TRUE, all.y = TRUE)

#populate countries missing from merged ones
ss$Country.x = ifelse(!is.na(ss$Country.y), ss$Country.y, ss$Country.x)

#Filter out NAs still remaining: one can't do anything about it
x = ss %>% filter(is.na(Country.x))  #17 rows remain: drop them

ss = ss %>% filter(!is.na(Country.x)) %>% select(-Country.y) %>% distinct()

rm(x)
#change column name
colnames(ss)[colnames(ss) == "Country.x"] = "Country"

#---------------PART 3: validate the data--------------------
#what could be validations:
#same category-subcategory pairs
#consistent customer ID-customer Name pairs
#--------------------------------------------------------------

#check if different cust id - cust name pairs exist
cust = ss %>% select(Customer.ID, Customer.Name) %>% distinct() %>%
       mutate(comb = paste(Customer.ID, Customer.Name, sep = "-")) %>%
       group_by(Customer.ID) %>% mutate(ct = n()) #4870 combinations exist
max(cust$ct) #each id appears exactly ones, no discrepancy

rm(cust)
#check if different category-subcategory pairs exist
cat = ss %>% select(Category, Sub.Category) %>% distinct() %>%
  mutate(comb = paste(Category, Sub.Category, sep = "-")) %>%
  group_by(Sub.Category) %>% mutate(ct = n()) #4870 combinations exist
max(cat$ct) #Some subcategories appear twice, examine by filtering these

cat_f = cat %>% filter(ct > 1)
cat_f  #all 9 categories under office supplies appear under office supplies & stationary
#change office supplies and stationary to office supplies for uniformity
#gsub may not work here bcecause of duplication of names; use ifelse
ss$Category = ifelse(ss$Category == "Office Supplies & Stationary", 
                     "Office Supplies", ss$Category)
unique(ss$Category)

rm(cat, cat_f)
#---------------Part 4: start wrangling the data---------------------
#Insert new columns needed for analysis

#get unit price as Sales / Quantity
ss$UnitPrice = ss$Sales/ss$Quantity

#create a flag for disocunted orders
ss = ss %>% mutate(Disc_flag = ifelse(Discount > 0, 1, 0))

#get month/year for each order date and also get days from order to shipment
ss = ss %>% mutate(month_year = format(as.Date(Order.Date), '%Y-%m')) %>%
     mutate(OrderToShip = difftime(Ship.Date, Order.Date, units = 'days'))

#create a flag if customer (cust ID) has made more than 8 
#orders (order id) in 4 years, classify as loyal
#first get cust id - order id combination, track unique entries
#then take count and classify
track = ss
track = track %>% select(Customer.ID, Order.ID) %>% distinct() %>%
        group_by(Customer.ID) %>% mutate(ct = n()) %>%
        mutate(cust_class = ifelse(ct > 8, "Loyal", "Regular")) %>%
        ungroup()

#drop order id, ct columns and take unique at customer id lever
track = track %>% select(-Order.ID, -ct) %>% distinct()
table(track$cust_class)

#merge track with ss
ss = merge(ss, track, by = "Customer.ID", all.x = TRUE)
rm(track)

#calculate profit margin as Profits / (Sales - Profits)
ss$profit_margin = round(ss$Profit / (ss$Sales - ss$Profit), digits = 2)

#check if row id appears exactly once and reorder columns
x = data.frame(table(ss$Row.ID)) #it's unique
rm(x)

colnames(ss)
ss_final = ss %>% select(16, 9, 1, 5, 29, 2:3, 22, 15, 7, 24,
                  11:12, 4, 23, 25, 14, 17, 13, 30, 21,
                  6, 26, 8, 19, 28, 27, 10, 20, 18)

write_xlsx(ss_final, "superstore_final.xlsx")


#-----------------PART 5: Aanalyze the data-----------------
