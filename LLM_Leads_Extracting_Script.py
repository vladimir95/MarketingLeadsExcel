from pprint import pprint
import os, sys, csv, xlrd, time, datetime, xlwt, re
from dateutil import parser

# @author - https://www.linkedin.com/in/vladimir-martintsov/

# PYTHON SCRIPT FOR COLLECTING THE RELEVANT INFORMATION FORM THE SOURCE DOUCMENTS, CLEANING IT, AND EXPORTING INTO SINGLE XLS FILE FOR FURTHER PROCESSING IN MS ACCESS
# The development process was aimed to have one source code file, where the client can go ahead and modify certain values, such as document column names, in order to 
# adapt to the changes of the format of the exported source documents. This gives flexibility to the marketing team to modify the column names in case if their sources export formats change. 

# The marketing team draws the relevant information from the relevant columns of the raw source documents stored in CSV and XLS formats. Usually CSV sources will come from Unbounce marketing
# platform and XLS - from MachForms marketing platform. Hence the naming of the global variables - the column names - for each of the platforms, as the column names differ between the two.
# The methods csv_opener() and xls_opener() take care of the importing raw data from CSV and XLS source documents and doing basic cleaning such as removing duplicates (through duplicate_checker() method).

# The client's specification is the data from XLS sources comes clean and does not need extra cleaning, unlike CSV sources; csv_data_cleaner() method takes care of that as well. 

# In the end, once we extracted all the relevant information from both sources and cleaned it, we bring it to the same format (export_source_leads() method), and export the data in a new XLS document.
# That XLS will be used to import the data into MS Access database, and should have clean data from all sources altogether.  

# In order to better understand the flow of the program, follow the main() method. We start by scanning the /Source directory and extracting information from CSV and XLS files present. Then, we clean the CSV data,
# and export everything to the external XLS named as per definition of export_xls_filename variable below.


#-------------------------------------------------------------------------------------------------------------------------------------------------------
# GLOBALLY DEFINED INFORMATION - SOURCE DOCUMENT COLUMN NAMES, LEADS EMAILS AND NAMES, AND OTHER RELEVANT INFO.

#-------------------------------------------------------------------------------------------------------------------------------------------------------
# ADD THE POSSIBLE MARKETING LEAD'S NAME HERE. NOTE: WHEN A RECORD WITH THIS NAME IS FOUND IN SOURCE FILE - THIS RECORD WILL BE CONSIDERED AS "WEIRD".
# Weird records will be exported separately for manual verification
osgoode_leads = ["stewart laszlo", "laszlo stewart",
				 "test ", "asd ", "test test", "asd asd", 
				 "farzana crocco", "crocco farzana", " ", 
				 "patricia pazos", "andrea chau", 
				 "sarah alexander", "sam butt"]

# ADD THE POSSIBLE MARKETING LEAD'S EMAIL HERE. NOTE: WHEN A RECORD WITH THIS EMAIL IS FOUND IN SOURCE FILE - THIS RECORD WILL BE MARKED AS "WEIRD".
# Weird records will be exported separately for manual verification
osgoode_leads_emails = ["achau@osgoode.yorku.ca","sambutt@yorku.ca",
						"sambutt@osgoode.yorku.ca","rbahrami@yorku.ca", 
						"rbahrami@osgoode.yorku.ca","ttendean@osgoode.yorku.ca",
						"ttendean@yorku.ca","slaszlo@osgoode.yorku.ca",
						"jmelancon@osgoode.yorku.ca",
						"eshallerauslander@osgoode.yorku.ca",
						"ppazos@osgoode.yorku.ca", "salexander@osgoode.yorku.ca",
						 "speakto@gmail.com", "chaumein@gmail.com",
						  "speaktosam@gmail.com", "test@gmail.com", 
						  "pdabek@gmail.com"]
#-------------------------------------------------------------------------------------------------------------------------------------------------------
# CSV COLUMNS FROM UNBOUNCE WEBSITE. 

# WHEN YOU EXPORT THE CSV, MAKE SURE THAT RELEVANT COLUMN NAMES ARE NAMED PROPERLY. CURRENTLY, THE 
# RELEVANT COLUMN NAMES ARE INDICATED AS VARIABLE NAMES HERE. PLEASE MODIFY THEM AS PER COLUMNS NAMES THAT UNBOUNCE GENERATES WHEN EXPORTING THE LEADS.
# If the script will not be able to find the relevant column, the error will be shown.

unbounce_date_submitted_column = "date_submitted"
unbounce_time_submitted_column = "time_submitted"
unbounce_first_name_column = "first_name"
unbounce_last_name_column = "last_name"
unbounce_email_column = "email"
unbounce_how_you_found_about_us = "utm_source"
unbounce_type_of_lead_column = "utm_medium"
unbounce_specialization_column = "page_variant_name"
unbounce_jd_llb_column = "do_you_have_a_jdllb"
unbounce_phone = "Phone Number"
unbounce_city = "city"
unbounce_province = "stateprovince"
unbounce_country = "country"

# The "specialization 2" information is not populated from source Unbounce document, as it does not exist in there in first place. We leave this column
# as is and do not collect any data as part of this script. Instead, we will populate the relevant data structures with an empty string for the Specialization field

# For coders: you can insert information from this column in csv_opener() method, if the specialization 2 column appears in CSV. At the time of coding this script,
# this column was not present, so I kept it as empty and unused
unbounce_specialization2_column = ""

#-------------------------------------------------------------------------------------------------------------------------------------------------------
# XLS COLUMNS FROM MACHFORMS WEBSITE. 
# Same story as the Unbounce. If the script will not be able to find the column, an error will be thrown. Please make sure the XLS column names are up to date.

machform_date_column = "Date Created"
machform_first_name_column = "Name - First"
machform_last_name_column = "Name - Last"
machform_email_column = "Email"
machform_specialization_column1 = "Program of Interest 1"
machform_specialization_column2 = "Program of Interest 2"
machform_how_you_found_about_us1 = "How did you first learn about the program?"
machform_type_of_lead_column = "Type of Lead"
machform_jd_llb_column = "Do you have a JD/LLB?"
machform_phone = "Phone Number"
machform_city = "City"
machform_province = "State/Province/Region"
machform_country = "Country"
machform_recieve_updates = "I would like to receive updates about my program of interest from Osgoode Professional Development. -- "

#machform_how_you_found_about_us2 = "How did you first "+
						#"learn about the Part-time Professional LLM?"


#-------------------------------------------------------------------------------------------------------------------------------------------------------
# THE LIST WITH ALL THE PROGRAMS AND THEIR KEYWORDS USED FOR CLEANING THE DIRTY DATA FROM UNBOUNCE SOURCE CSVs.

# The keywords in dirty specialization 1 data may show exactly which proper Specialization 1 is meant for this entry. 
# Note: The dirty data comes in lowered letters due to proper comparison process used in csv_opener() method

# Once we detect the keyword in dirty Specialization 1, we will immediately switch it to the relevant clean version of Specialziation 1.

# IF YOU WANT TO ADD A KEYWORD AND A CLEAN SPECIALIZATION MAKE SURE THAT THE KEYWORD ITSELF IS ONLY IN SMALL LETTERS! 

program_keywords = {
							"admin" : "Administrative Law",
							"bank" : "Banking and Financial Services Law",
							"canad" : "Canadian Common Law",
							"cert" : "Certificate in Foundations for Graduate Legal Studies",
							"consti" : "Constitutional Law",
							"crimi" : "Criminal Law and Procedure",
							"disp" : "Dispute Resolution",
							"ener" : "Energy and Infrastructure Law",
							"enfor" : "Gradute Diploma in Law for Law Enforcement",
							"heal" : "Health Law",
							"intel" : "Intellectual Property Law",
							"labo" : "Labour Relations and Employment Law",
							"onlin" : "Online NCA Exam Prep",
							"priva" : "Privacy and Cybersecurity Law",
							"securities" : "Securities Law",
							"yuel" : "YUELI's Intensive Advanced Legal English Program"
							}

# There are a bit more complicated cases of overlapping keywords. Only the combination of these keywords will properly determine the proper clean Specialization. Unfortunately, for proper keyword
# detection of the Specializations that may appear in the future, you may require more development work. We explicitly define the overlapping keywords here and use them logically in csv_data_cleaner() method. 

# The workaround with overlapping keywords is to come up with unique keywords that charecterize a certain program and use it internally in Unbounce when creating page variants, as Specialization 1 
# is drawn from page_variant_name, which is created internally by marketing team. It is easier to agree internally that a custom new keyword is meant for the new Specialization that was introduced after this script was developed
business_key = "busi"
international_key = "interna"
full_time_key = "full"
llm_key = "llm"
professional_key = "profess"
program_key = "prog"
part_time_key = "part"
tax_key = "tax"
study_portal_key = "study"
single_course_key = "enrol"

business_law = "Business Law"
international_business_law = "International Business Law"
full_time_law = "Full-time LLM"
single_course_pt = "Single Course Enrollment - PT"
single_course_ft = "Single Course Enrollment - FT"
international_programs_law = "International Programs"
part_time_law = "Part-time LLM"
tax_ft_law = "Tax Law - FT"
tax_pt_law = "Tax Law - PT"


#-------------------------------------------------------------------------------------------------------------------------------------------------------
# THE LIST WITH ALL THE UNBOUNCE TYPE OF LEAD AND THEIR KEYWORDS USED FOR CLEANING THE DIRTY DATA FROM UNBOUNCE SOURCE CSVs.

# The keywords in dirty Type of Lead data may show exactly which proper Type of Lead is meant for this entry. 
# Note: The dirty data comes in lowered letters from Unbounce, as it was also lowered in csv_opener method.

# Once we detect the keyword in dirty Type of Lead, we will immediately switch it to the relevant clean version of Type of Lead. 

# NOTE: IF THE TYPE OF LEAD IS EMPTY, THE SCRIPT WILL MARK THAT ENTRY AS "WEIRD", AND EXPORT ACCORDINGLY

# IF YOU WANT TO ADD A KEYWORD AND A CLEAN TYPE OF LEAD MAKE SURE THAT THE KEYWORD IS ONLY SMALL LETTERS! 

type_of_lead_keywords = {
							"aggregate" : "Aggregate",
							"cpc" : "JMH",
							"digital" : "JMH",
							"email" : "Email",
							"enewsletter" : "Email",
							"events" : "Events",
							"link-posts" : "TBD: link-posts", # Link posts are TBD, not sure which category yet
							"outdoor" : "Ads",
							"print" : "Ads",
							"social+advertising" : "JMH",
							"sponsored" : "JMH",
							"video" : "Video",
							"web+ad" : "Ads",
							}



#-------------------------------------------------------------------------------------------------------------------------------------------------------
# ALL RELEVANT OUTPUT LEADS XLS EXPORT COLUMN AND FILE NAMES. 
# When we export the leads from both XLS and CSV sources, we organize the destination file with the following column names
# Order matters!

# IF YOU CHANGE THE ORDER OF THE COLUMNS (ITS NUMBER), YOU HAVE TO REARRANGE THEIR ORDER IN THIS LIST BELOW AS WELL!

# IF YOU CHANGE THE NAME OF THE COLUMN, MAKE SURE TO REFLECT YOUR CHANGES IN EXPORT_SOURCE_LEADS() METHOD! SCROLL DOWN TO THE METHOD TO SEE MORE!

export_column_names_list = ["Date Added", "First Name", "Last Name", "Email", 
							"Type of Lead", "Specialization 1", "Specialization 2", 
							"How did you hear about us?", 
							"JD/LLB?", "Phone", "City", "Prov", "Country", "Notes"]

export_column_names = {	# Column name     			Column Number
						"Date Added" : 					0,
						"First Name" : 					1,
						"Last Name" : 					2,
						"Email" : 						3,
						"Type of Lead" : 				4,
						"Specialization 1" : 			5,
						"Specialization 2" : 			6,
						"How did you hear about us?" : 	7,
						"JD/LLB?" : 					8,
						"Phone" :						9,
						"City" : 						10,
						"Prov" : 						11,
						"Country" : 					12,
						"Source" :		 				13,
						}

# HERE YOU CAN ASSIGN A FILE NAME FOR THE XLS FILE THAT YOU WILL BE EXPORTING THE CLEAN LEADS INTO AS WELL AS THE SHEET NAMES THAT WILL BE CREATED INSIDE THAT FILE

export_xls_filename = "All_source_leads.xls"
export_leads_sheet_name = "All source leads"
export_weird_leads_sheet_name = "All weird leads"
#---------------------------------------------------------------------------------------------------------------------------------------------------------
# Now for the nerdy stuff. Please do not modify anything below this line!
#---------------------------------------------------------------------------------------------------------------------------------------------------------

# All the entries - CSV, XLS, duplicates, and weird entries - will have a running ID (as dictionary keys) for the internal use in their respective dictionaries.


# All information from all CSVs will be collected into all_csv_entries_info dictionary and will have running total, weird and duplicate counters used for the stats.
csv_entry_id = 1
all_csv_entries_info = {}

csv_total_entries_count = 0
csv_total_duplicate_count = 0
csv_weird_leads_total_count = 0


# All information from all XLSs will be collected into all_xls_entries_info dictionary and will have running total, weird and duplicate counters used for the stats.
# We also have a counter for the people who did not check "I want to receive email updates..." checkbox for the stats.
xls_entry_id = 1
all_xls_entries_info = {}

xls_total_entries_count = 0
xls_total_duplicate_count = 0
xls_weird_leads_total_count = 0
xls_unchecked_count = 0


# All weird leads that were extracted from XLS and CSVs altogether will be stored here for further manual processing
# The weird leads will have their own id number
all_weird_leads_id = 1
all_weird_leads_info = {}

# Duplicate count
all_duplicates_id = 1
all_duplicates = {}


def email_invaild(email):
	# Basic sanity check method for the email addresses. Check if only one @ is present, and the address does not have spaces or commas
	# @email - the email we are doing the sanity check for

	at_symbol_count = email.count("@")

	if " " not in email and "," not in email and at_symbol_count == 1:
		return False
	else:
		return True

def dict_tuple_updater(data, data_key, tuple_index, new_value):
	# This helper method is designed to update the value of any tuple entry in any dictionary.
	# @data - the dicitionary where we need to update the value tuple
	# @data_key - the key, whose value tuple we have to update
	# @tuple_index - the index of the value within the tuple that we need to update 
	# @new_value - the new value to be inserted instead of previous one at that tuple index

	new_list = list(data[data_key])
	new_list[tuple_index] = new_value
	data[data_key] = tuple(new_list)

def duplicate_checker(data, flag):
	# Duplicate checker for entries from the source files - used in csv_opener and xls_opener methods.
	# We compare every entry to all other entries in the information we read from given CSV/XLS file. We do this comparison withing each file only, thus removing duplicates.
	# If there is match between the:
	#
	# - email (logical)and program - that is, the name may differentiate - the person made a mistake in name spelling
	# - name (logical) and program - that is, the email may differentiate - the person made a mistake in email spelling
	#
	# Then, we compare the timestamps of the two records and see which one is the latest. 
	# The Customer's requirements are: 
	# We do that as we want to remove only people who expressed interest in the same program several times.
	# It makes sense, as the person filling out the submission form would submit the last form most accurately.

	# @data - the dictionary storing entries from one CSV/XLS files that has to be de-duplicated   
	# @flag - a string identifier to distinguish CSV or XLS files 

	global all_duplicates_id
	global all_duplicates
	global csv_entry_id
	global xls_entry_id
	global xls_total_duplicate_count
	global csv_total_duplicate_count

	# Thus, we preserve the very last(in time) record and delete this entry from the dictionary
	for comparison_id in list(data):
		try:
			comparison_email = data[comparison_id][2]
			comparison_name = data[comparison_id][1]
			comparison_specialization = data[comparison_id][4]
			comparison_date = data[comparison_id][0]
		except KeyError:
			continue


		for entry_id in list(data):
			try:
				entry_email = data[entry_id][2]
				entry_name = data[entry_id][1]
				entry_specialization = data[entry_id][4]
				entry_date = data[entry_id][0]

			# If the key was already removed, we move to the next key in the list of keys
			except KeyError:
				continue
			# Case 1: Found an email and program match
			if (comparison_email == entry_email) and (comparison_specialization == entry_specialization):
				# Now compare the dates and keep the record with the latest date

				# We are looking at the same record
				if(comparison_date == entry_date):
					continue

				# The comparison record is later than the one we are currently iterating through,
				# we delete the current entry
				# That way we keep the "latest" record only
				elif(comparison_date > entry_date):
					# Store this duplicate information; in case if we need to check what the duplicates are
					all_duplicates[all_duplicates_id] = data[entry_id]
					all_duplicates_id +=1

					# Increment the respective counters for statistics
					if (flag == "csv"):
						csv_total_duplicate_count +=1
					elif(flag == "xls"):
						xls_total_duplicate_count +=1
					# we delete the duplicate item from the main data dictionary
					del data[entry_id]
					continue

			# By now the people with Duplicate emails (AND duplicate names, respectively) are removed
			# Case 2: Found a name and program match
			if (comparison_name == entry_name) and (comparison_specialization == entry_specialization):
				#Now we need to compare the dates and keep the record with the lastest date

				# We are looking at the same record
				if(comparison_date == entry_date):
					continue

				# The comparison record is later than the one we are currently iterating through,
				# we delete the current entry 
				# That way we keep the "latest" record only
				elif(comparison_date > entry_date):
					# Store this duplicate information; in case if we need to check what the duplicates are
					
					all_duplicates[all_duplicates_id] = data[entry_id]
					all_duplicates_id +=1

					# Increment the respective counters for statistics
					if (flag == "csv"):
						csv_total_duplicate_count +=1
					elif(flag == "xls"):
						xls_total_duplicate_count +=1
						
					# we delete the duplicate item from the main data dictionary
					del data[entry_id]
					continue

def csv_opener(filename):
	# The method that reads ONE CSV source file that is taken from Unbounce or any other source. The other source CSV export format must match the Unbounce's format. The column names are defined globally to reflect any changes if needed. 
	# Only the relevant data will be extracted, de-duplicated and then stored in the global dictionary (all_csv_entries_info) of all leads from Unbounc. all_csv_entries_info is indexed with csv_entry_id 
	# The "weird" leads - with abnormal first and last names - will be stored separately for further manual processing by marketing team. 

	# @filename - the CSV file from Unbounce to be processed


	global csv_entry_id
	global csv_total_entries_count
	global all_csv_entries_info
	global csv_weird_leads_total_count
	global osgoode_leads
	global osgoode_leads_emails
	global all_weird_leads_id
	global all_weird_leads_info

	global unbounce_date_submitted_column
	global unbounce_time_submitted_column
	global unbounce_first_name_column
	global unbounce_last_name_column
	global unbounce_email_column
	global unbounce_how_you_found_about_us
	global unbounce_type_of_lead_column
	global unbounce_specialization_column
	global unbounce_jd_llb_column
	global unbounce_phone
	global unbounce_city
	global unbounce_province
	global unbounce_country

	# Import the Specialization 2 column name from Unbounce, in case if it appears in the future
	#global unbounce_specialization2_column

	current_source_csv = {}


	with open(filename, "r", encoding="ISO-8859-1") as csv_file:
		reader = csv.DictReader(csv_file, delimiter=",")
		for line in reader:
			try:

				#Got an entry - count it in
				csv_total_entries_count +=1

				#get the date
				timestamp = line[unbounce_date_submitted_column] + " " + line[unbounce_time_submitted_column]
				timestamp = timestamp[:-3]
				try:

					# parse the date
					date = parser.parse(timestamp)
				except:
					# The data might come in super dirty, so we mark this entry as weird and let the marketing team take care of it, if anything. 
					all_weird_leads_info[all_weird_leads_id] = date_value, full_name_value, email_value, type_of_lead_value, specialization1_value, specialization2_value, how_you_found_about_us_value, jd_llb_value, phone_value, city_value, province_value, country_value, "Unable to parse date"
					all_weird_leads_id +=1
					xls_weird_leads_total_count +=1
					continue

				# We gather the full name for comprehensive duplicate comparison by duplicate_checker() method
				# We make sure to remove all extra white spaces people could put by accident
				# We also lower the letters in both names, email and specialization for proper comparison, as some people may miss capitals or make typing errors. We also strip the white spaces.
				# We do not lower any other fields as they are not used in the comparisons. 
				# We lower and dtrip the type of lead, as it will be cleaned later in csv_data_cleaner() method

				name = line[unbounce_first_name_column].lower().strip() + "\t" + line[unbounce_last_name_column].lower().strip()
				# Email will be used by duplicate_checker()
				email = line[unbounce_email_column].lower().strip()
				how_you_found_about_us = line[unbounce_how_you_found_about_us]
				type_of_lead = line[unbounce_type_of_lead_column].lower().strip()
				# Specialization will be used by duplicate_checker()
				specialization = line[unbounce_specialization_column].lower()#.strip()
				jd_llb = line[unbounce_jd_llb_column]
				city = line[unbounce_city]
				province = line[unbounce_province]
				country = line[unbounce_country]

				#phone = line[unbounce_phone]
				phone = ""

				#specialization2 = line[unbounce_specialization2_column]
				specialization2 = ""




				# Quick clean of City and Province fields:
				# - fully capitalize 2 character words in Province
				# - lowercase all the words and capitalize first letters of all words that are longer than 2 characters in City and Province

				try:
					#city = re.sub('[^A-Za-z0-9\.' ']+', '', city)
					#province = re.sub('[^A-Za-z0-9\.' ']+', '', province)

					city = city.lower().title()

					# If province/state has 2 characters, capitalize them as abbreviations (i.e. ON, BC, AB, etc. ), and remove all accidental white spaces
					if(len(province)==2):
						province = province.strip().upper()
					else:	
						province = province.lower().title()
				except:
						
					# The data might come in super dirty, so we mark this entry as weird and let the marketing team take care of it, if anything. 
					all_weird_leads_info[all_weird_leads_id] = date_value, full_name_value, email_value, type_of_lead_value, specialization1_value, specialization2_value, how_you_found_about_us_value, jd_llb_value, phone_value, city_value, province_value, country_value, "City or Province are weird"
					all_weird_leads_id +=1
					xls_weird_leads_total_count +=1
					continue

				# Check if the entry has weird or test email from Osgoode leads who were testing the page and submitted the forms to see if the page works
				# We omit the entries with too short of (full) names (i.e. asdads), names that match osgoode leads names or their emails
				# We also consider the records with wrong emails as weird and will export them for manual verification
				if ((len(name) <=7) or (name in osgoode_leads) or (email in osgoode_leads_emails) or email_invaild(email)):

					#Populate the weird leads dictionary with relevant data and the filename it came from
					all_weird_leads_info[all_weird_leads_id] = date, name, email, type_of_lead, specialization, specialization2, how_you_found_about_us, jd_llb, phone, city, province, country, csv_file.name
					all_weird_leads_id +=1
					csv_weird_leads_total_count +=1
					continue

				# Put the good data in the global csv dictionary with the relevant information we needed
				current_source_csv[csv_entry_id] = date, name, email, type_of_lead, specialization, specialization2, how_you_found_about_us, jd_llb, phone, city, province, country
				csv_entry_id +=1

			# Error checking for the colmn names. If the DictReader, which takes the first row as column names will not be able to find the coulmn name in the first row as we defined before, it will throw a KeyError
			except KeyError as e:
				print("There was an error that encountered when reading the column names in the following file: " + str(filename.name))
				print("The script was unable to find the column named as " + str(e) + " in this file")
				print("Please refer to lines 24 -35 in the script and verify the column names for Unbounce CSV source documents")
				sys.exit(0)

			# The other exception may come from parsing date from the sheet. This has to be manually fixed before teh CSV being fed to the script. We use date and time to compare records down the road
			except ValueError as e:
				print("There was an error that encountered when reading the column names in the following file: " + str(filename.name))
				print("The script was unable to parse Date value in the Date and Time columns presumed to be in this file")
				print("Please refer to the lines 24 and 25 of this script to specify the respective Date and Time column names in the Unbounce Source files")
				sys.exit(0)

	# Once we are done with iterating through this CSV, close it to release the resources from memory.
	csv_file.close()

	# Now that we removed the weird leads, we check which entries are duplciates
	duplicate_checker(current_source_csv, "csv")

	# Now that the duplicates were removed, we put them into the global dictionary which will be later compared to the admissions database
	# Since we were incrementing the csv_entry_id, the records in this dictionary will never overlap and update each other
	all_csv_entries_info.update(current_source_csv)

def xls_opener(filename):
	# The method that reads ONE XLS source file that is taken from Mockforms or any other source. The other source XLS export format must match the Machforms's format. The column names are defined globally to reflect any changes if needed. 
	# Only the relevant data form relevant columns will be extracted, de-duplicated and then stored in the global dictionary of all leads from Mockforms. 
	# The "weird" leads - with abnormal first and last names - will be stored separately for further manual processing by marketing team.
	global all_xls_entries_info
	global xls_entry_id
	global xls_total_entries_count
	global xls_weird_leads_total_count
	global xls_unchecked_count
	global osgoode_leads
	global osgoode_leads_emails
	global all_weird_leads_id
	global all_weird_leads_info

	global machform_first_name_column
	global machform_last_name_column
	global machform_email_column
	global machform_date_column
	global machform_program_column
	global machform_specialization_column1
	global machform_specialization_column2
	global machform_how_you_found_about_us1
	global machform_type_of_lead_column
	global machform_jd_llb_column
	global machform_phone
	global machform_city
	global machform_province
	global machform_country
	global machform_recieve_updates

	#global machform_type_of_lead_column

	#TO DO: Figure out the specialization 2 column for machforms - standardize the name of that second column
	#global machform_specialization_column2

	current_source_xls = {}

	xlscolumns = {}

	xlsworkbook = xlrd.open_workbook(filename)
	sheet = xlsworkbook.sheet_by_index(0)

	# Grab the columns from xls file
	for column in range(sheet.ncols):
		xlscolumns[sheet.cell_value(0,column)] = column

	try:
		email_column = xlscolumns[machform_email_column]
		first_name_column = xlscolumns[machform_first_name_column]
		last_name_column = xlscolumns[machform_last_name_column]
		date_column = xlscolumns[machform_date_column]
		specialization1_column = xlscolumns[machform_specialization_column1]
		specialization2_column = xlscolumns[machform_specialization_column2]
		how_you_found_about_us_column = xlscolumns[machform_how_you_found_about_us1]
		type_of_lead_column = xlscolumns[machform_type_of_lead_column]
		jd_llb_column = xlscolumns[machform_jd_llb_column]
		#phone_column = xlsxcolumns[phone]
		city_column = xlscolumns[machform_city]
		province_column = xlscolumns[machform_province]
		country_column = xlscolumns[machform_country]
		receive_updates_column = xlscolumns[machform_recieve_updates]
	except KeyError as e:
		print("There was an error that encountered when reading the column names in the following file: " + str(filename.name))
		print("The script was unable to find the column named as " + str(e) + " in this file")
		print("Please refer to lines 49 - 62 in the script and verify the column names for Mockforms XLS source documents")
		sys.exit(0)	

	# Now extract the First Name, Last Name and Email into the output
	for row in range(sheet.nrows):
		# skip the header
		if row == 0:
			continue

		# Got an entry, increment the counter
		xls_total_entries_count +=1
		
		receive_updates_value = sheet.cell_value(row, receive_updates_column).lower()

		# skip the folks who do not want to receive the updates
		if receive_updates_value == "":
			xls_unchecked_count +=1
			continue

		# Now only folks who want to receive an update below this line

		# Grab the relevant values from the sheet, and lower those that are needed for duplicate_checker()
		email_value = sheet.cell_value(row, email_column).lower().strip()

		
		# We gather the full name for comprehensive duplicate comparison by duplicate_checker() method
		# We make sure to remove all extra white spaces people could put by accident
		# We also lower the letters in both names, email for proper comparison, as some people may miss capitals or make typing errors
		# Note that XLS source documents come with clean Specialization data, so we do not need to lower it for comparison
		# We do not lower any other fields as they are not used in the comparisons. 
		first_name_value = sheet.cell_value(row, first_name_column).lower().strip()
		last_name_value = sheet.cell_value(row, last_name_column).lower().strip()
		full_name_value = first_name_value + "\t" + last_name_value
		specialization1_value = sheet.cell_value(row, specialization1_column)
		specialization2_value = sheet.cell_value(row, specialization2_column)
		how_you_found_about_us_value = sheet.cell_value(row, how_you_found_about_us_column)
		type_of_lead_value = sheet.cell_value(row, type_of_lead_column)
		jd_llb_value = sheet.cell_value(row, jd_llb_column)
		#phone_value = sheet.cell_value(row, phone_column)
		phone_value = ""
		city_value = sheet.cell_value(row, city_column)
		province_value = sheet.cell_value(row, province_column)
		country_value = sheet.cell_value(row, country_column)


		try:
			# Parse the date to be datetime object for later comparisons
			# This will take the string value and parse it into the date
			date_value = sheet.cell_value(row, date_column)
			date_value = parser.parse(date_value)
		except:
			try:

				# ADD COMMENTS AND CHECK FOR SHIT CASES HANDLING ALL THROUGHOUT!!!!
				# Since the data might come in dirty we might be working with Excel float numbers, which require a different solution

				date_value = sheet.cell_value(row, date_column)
				temp = datetime.datetime(1899, 12, 30)
				delta = datetime.timedelta(days=date_value)

				date_value = temp+delta
			except:	

				# Give up and let the date be fixed manually. Send the entry to the weird leads.
				all_weird_leads_info[all_weird_leads_id] = date_value, full_name_value, email_value, type_of_lead_value, specialization1_value, specialization2_value, how_you_found_about_us_value, jd_llb_value, phone_value, city_value, province_value, country_value, "Unable to parse the date"
				all_weird_leads_id +=1
				xls_weird_leads_total_count +=1
				continue



		# Quick clean of City and Province fields:
		# - fully capitalize 2 character words in Province
		# - lowercase all the words and capitalize first letters of all words that are longer than 2 characters in City and Province

		try:
			#city_value = re.sub('[^A-Za-z0-9\.' ']+', '', city_value)
			#province_value = re.sub('[^A-Za-z0-9\.' ']+', '', province_value)
			city_value = city_value.lower().title()
			# If province has 2 characters, capitalize them, and remove all accidental white spaces
			if(len(province_value)==2):
				province_value = province_value.strip().upper()
			else:	
				province_value = province_value.lower().title()

		except:

			# The data might come in super dirty, so we mark this entry as weird and let the marketing team take care of it, if anything. 
			all_weird_leads_info[all_weird_leads_id] = date_value, full_name_value, email_value, type_of_lead_value, specialization1_value, specialization2_value, how_you_found_about_us_value, jd_llb_value, phone_value, city_value, province_value, country_value, "City or Province are weird"
			all_weird_leads_id +=1
			xls_weird_leads_total_count +=1
			continue



		#Check if the entry has weird or test email from Osgoode leads who were testing the page and submitted the forms to see if the page works
		# We omit the entries with too short of names, names that match osgoode leads names or their emails
		# We also consider the records with wrong emails as weird and will export them for manual verification
		if ((len(full_name_value) <=7) or (full_name_value in osgoode_leads) or (email_value in osgoode_leads_emails) or email_invaild(email_value)):

			#Populate the weird leads dictionary with relevant data and the filename from where it came from
			all_weird_leads_info[all_weird_leads_id] = date_value, full_name_value, email_value, type_of_lead_value, specialization1_value, specialization2_value, how_you_found_about_us_value, jd_llb_value, phone_value, city_value, province_value, country_value, filename.name
			all_weird_leads_id +=1
			xls_weird_leads_total_count +=1
			continue
																					

		# Now prepare the dictionary with the relevant information we needed
		current_source_xls[xls_entry_id] = date_value, full_name_value, email_value, type_of_lead_value, specialization1_value, specialization2_value, how_you_found_about_us_value, jd_llb_value, phone_value, city_value, province_value, country_value
		xls_entry_id +=1
		
	# Now that we removed the weird leads, we check which entries are duplciates
	duplicate_checker(current_source_xls, "xls")

	# Now that the duplicates were removed, we put them into the global dictionary which will be later compared to the admissions database
	all_xls_entries_info.update(current_source_xls)

def csv_data_cleaner():

	# A method that cleans up the entire CSV source data before it is getting exported. Unbounce has dirty data.
	# Data from Mockforms(XLS) comes already clean and does not need additional cleaning

	# We have to clean the Type of Lead field - assign each abbreviation to the proper universal names that were agreed in marketing team
	# The records that have Type of Lead entry as empty - are considered as weird leads, and will be marked accordingly

	# We have to clean Specialization 1 field. Based on what keywords the specialization has, we assign a proper name to it for consistency
	# Note that we only work with Specialization 1. Specialization 2 comes blank from Unbounce
	# We are given a list of programs that people may enroll (program_keywords) and the specialization that was exported from Unbounce
	# The exported data is dirty, but all of it has the keywords, based on which we can replace the specialization entry with cleaner data

	global all_csv_entries_info
	global all_weird_leads_info
	global all_weird_leads_id
	global csv_weird_leads_total_count

	global type_of_lead_keywords
	global program_keywords

	global business_key
	global international_key
	global full_time_key
	global llm_key
	global professional_key
	global program_key
	global part_time_key
	global tax_key
	global study_portal_key
	global single_course_key

	global business_law
	global international_business_law
	global full_time_law
	global international_programs_law
	global part_time_law
	global tax_ft_law
	global tax_pt_law
	global single_course_ft
	global single_course_pt

	# Start cleaning the CSV data
	for entry_id in list(all_csv_entries_info):
		try:

			# Type of Lead is the 3rd entry in the value tuple of the dictionary
			# Specialization is 4th entry and the "How you heard about us" is 6th

			# Note: for Tax Law, if "How you heard about us" column for a Tax Law Specialization has "study+portals", that means that Tax Law is Full Time
			# In all other cases, the Tax Law is advertised as Part Time on source pages
			# We do not use data from "How you heard about us" column for other types of Specialization; we only use it to distinguish Tax Law in CSV sources from Unbounce

			# The dirty Type of Lead data from all CSVs
			type_of_lead = all_csv_entries_info[entry_id][3]
			# The dirty Specialization 1 data from all CSVs 
			spec = all_csv_entries_info[entry_id][4]
			# The "How you heard about us" data from all CSVs
			heard = all_csv_entries_info[entry_id][6]

			# The country comes into play when distinguishing the Single Course Enrollment
			country = all_csv_entries_info[entry_id][10].lower().strip()

			# Check the type of lead keywords and see if the dirty Type of Lead can be translated into a meaningful Type of Lead. We are doing strict comparison; EVEN if the value is blank, we will mark them weird.
			# We stripped the whitespaces and lowered the letters upon import in csv_opener() method, and the data does not come as dirty as with Specialization. So we should be able to do a direct translation
			for keyword in type_of_lead_keywords.keys():
				if (keyword in type_of_lead):
					# We found the keyword in type of lead that matches the type of lead that was agreed by the marketing team, so we update it accordingly
					dict_tuple_updater(all_csv_entries_info, entry_id, 3, type_of_lead_keywords[keyword])
					break
			# Else we have to mark this lead as Weird, and delete it from the pool of all the entries. Note: even blanks are marked as weird.
			else: 
				#print("Found a weird lead!")
				# Add this entry to the weird leads dictionary

				# Instead of putting the filename, we put a message that the Specialization 1 is weird
				entry_info = "Type of Lead is weird"
				
				# We want to add "Type of lead is weird" string to the end of the tuple for this weird entry. We did that before when we would add Source file at the end of the weird entry.
				# This time we add another distinctive message for marketing team to filter for after.

				# Append the "Type of Lead is weird" to the tuple in all_csv_entry_info[entry_id], which presumably contains weird Type of Lead, by making a new list first
				appended_info = list(all_csv_entries_info[entry_id])
				appended_info.append(entry_info)

				# Capitalize the Type of Lead before exporting it 
				appended_info[3] = appended_info[3].title()

				# Now add the updated tuple to the weird leads
				all_weird_leads_info[all_weird_leads_id] = tuple(appended_info)
				all_weird_leads_id +=1
				csv_weird_leads_total_count +=1

				# remove the entry from CSV entries
				del all_csv_entries_info[entry_id]

				# We logically continue to the next iteration
				continue


			# Now we move on to cleaning the Specialization 1

			# Check program keywords and see if the dirty Specialization 1 comes with any of the keywords. If yes, we update the all_csv_entries_info dictionary's tuple with the clean Specialization 1 name.
			for keyword in program_keywords.keys():
				if (keyword in spec):
					# We found the keyword in dirty spec, so we update the data with the clean Specialization name
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, program_keywords[keyword])
					break
			# Else, we look at the specific cases of overlapping keywords, and we cannot find any - we mark them as weird
			else:

				# We distinguish International vs Business Law
				if((business_key in spec) and (international_key not in spec)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, business_law)
				elif((business_key in spec) and (international_key in spec)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, international_business_law)

				# International LLM is considered as Full-time LLM as well, so we combine the two cases
				elif(((international_key in spec) and (llm_key in spec)) or (full_time_key in spec)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, full_time_law)

				# International Program must be found by two keywords
				elif((international_key in spec) and (program_key in spec)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, international_programs_law)

				# Professional LLM is considered as Part-time LLM as well, so we combine the two cases
				elif(((professional_key in spec) and (llm_key in spec)) or (part_time_key in spec)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, part_time_law)

				# Customer requirement: if heard = study+portals, then full time, otherwise - part time	
				elif((tax_key in spec) and (study_portal_key in heard)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, tax_ft_law)
				elif((tax_key in spec) and (study_portal_key not in heard)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, tax_pt_law)

				# Single course enrollment is country based. If the person is not from Canada - full-time. Else, we default it to part-time (the country is Canada or empty)
				elif((single_course_key in spec) and (len(country) > 1) and ("canad" not in country)):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, single_course_ft)
				elif((single_course_key in spec) and ((len(country) < 1) or ("canad" in country))):
					dict_tuple_updater(all_csv_entries_info, entry_id, 4, single_course_pt)

				# We skip empty Specializations and leave them as is
				elif(len(spec.strip()) == 0):
					continue

				# We are unable to figure out what program this person is interested in, so we put them into weird list for further manual processing
				else:

					#print("Found a weird lead!")
					# Add this entry to the weird leads dictionary

					# Instead of putting the filename, we put a message that the Specialization 1 is weird
					entry_info = "Specialization 1 is weird"
					
					# We want to add "Specialization 1 is weird" string to the end of the tuple for this weird entry. We did that before when we would add Source file at the end of the weird entry.
					# This time we add another distinctive message for marketing team to filter on after.

					# Append the "Specialization 1 is weird" to the tuple in all_csv_entry_info[entry_id], which presumably contains weird Specialization 1, by making a new list first

					appended_info = list(all_csv_entries_info[entry_id])
					appended_info.append(entry_info)

					# Capitalize the Specialization 1 before exporting it 
					appended_info[4] = appended_info[4].title()

					all_weird_leads_info[all_weird_leads_id] = tuple(appended_info)
					all_weird_leads_id +=1
					csv_weird_leads_total_count +=1

					# remove the entry from CSV entries
					del all_csv_entries_info[entry_id]
		# Since we may delete an entry from all_csv_entries_info, we skip the deleted key just in case			
		except KeyError:
			continue

def spec_updater(entry_id,spec,spec_index,heard, country):
	# This helper method for weird_data_cleaner() method grabs the specialization and updates it as per program_keywords. If there is no match, it leaves it as is.
	# @entry_id - the key id of the dictionary key-value pair
	# @spec - the Specialization to be verified
	# @spec_index - the index of the Specialization within the value tuple for the entry_id key
	# @heard - the entry in "how did you hear about us?" that is used for the Tax Law
	# @country - the entry in "country" field that is used for the Single Course Enrollment 
	global program_keywords
	global all_weird_leads_info

	global business_key
	global international_key
	global full_time_key
	global llm_key
	global professional_key
	global program_key
	global part_time_key
	global tax_key
	global study_portal_key
	global single_course_key

	global business_law
	global international_business_law
	global full_time_law
	global international_programs_law
	global part_time_law
	global tax_ft_law
	global tax_pt_law
	global single_course_pt
	global single_course_ft

	# If the Specialization comes empty - we are not going to bother with it and return
	if (spec == ""):
		return

	# Check program keywords and see if the dirty Specialization 1 comes with any of the keywords. If yes, we update the all_csv_entries_info dictionary's tuple with the clean Specialization 1 name.
	for keyword in program_keywords.keys():
		if (keyword in spec):
			# We found the keyword in dirty spec, so we update the data with the clean Specialization name
			dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, program_keywords[keyword])
			break
		# Else, we look at the specific cases of overlapping keywords, and we cannot find any - we mark them as weird
		else:

			# We distinguish International vs Business Law
			if((business_key in spec) and (international_key not in spec)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, business_law)
			elif((business_key in spec) and (international_key in spec)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, international_business_law)

			# International LLM is considered as Full-time LLM as well, so we combine the two cases
			elif(((international_key in spec) and (llm_key in spec)) or (full_time_key in spec)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, full_time_law)

			# International Program must be found by two keywords
			elif((international_key in spec) and (program_key in spec)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, international_programs_law)

			# Professional LLM is considered as Part-time LLM as well, so we combine the two cases
			elif(((professional_key in spec) and (llm_key in spec)) or (part_time_key in spec)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, part_time_law)

			# Customer requirement: if heard = study+portals, then full time, otherwise - part time	
			elif((tax_key in spec) and (study_portal_key in heard)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, tax_ft_law)
			elif((tax_key in spec) and (study_portal_key not in heard)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, tax_pt_law)
			# Single course enrollment is country based. If the person is not from Canada - full-time. Else, we default it to part-time (the country is Canada or empty)
			elif((single_course_key in spec) and (len(country) > 1) and ("canad" not in country)):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, single_course_ft)
			elif((single_course_key in spec) and ((len(country) < 1) or ("canad" in country))):
				dict_tuple_updater(all_weird_leads_info, entry_id, spec_index, single_course_pt)

			# We keep the Specialization as is
			else:
				continue

def weird_data_cleaner():
	# This method cleans the Weird data in all_weird_leads_info dictionary before exporting it. The Marketing team will manually check this list after exporting it, but cleaning up the Specializations and Type of Lead(if applicable)
	# will be annoying. So we do this here, prior exporting weird leads into XLS. At this point, all the weird leads should be collected in all_weird_leads_info dictionary, and we can clean them.

	#NOTE: We can only clean the cases that are known to us - similarly to csv_data_cleaner() method, using the data in type_of_lead_keywords and program_keywords data structures. If we encounter any other entry in there - we will simply
	# skip it and let the marketing team take care of it. 

	# We clean Specialization 1 and 2 only for Unbounce entries - we distinguish them from MachForms by them being either empty, or having a keyword from type_of_lead_keywords data structure

	global all_weird_leads_info
	global type_of_lead_keywords


	# Start going over the weird leads
	for weird_lead_id in all_weird_leads_info.keys():

		# Type of Lead is stored in column 3
		type_of_lead = all_weird_leads_info[weird_lead_id][3]

		# Specialization 1 is stored in column 4
		spec1 = all_weird_leads_info[weird_lead_id][4]

		# Specialization 2 is stored in column 5
		spec2 = all_weird_leads_info[weird_lead_id][5]

		# For Tax Law, if the "How did you hear about us?" is "study+portals", then it is Tax Law - FT
		heard = all_weird_leads_info[weird_lead_id][6]

		# The country comes into play when distinguishing the Single Course Enrollment
		# For Single Course Enrollment, if the Country is not Canada - FT. If country is Canada OR empty - then we default it to PT
		country = all_weird_leads_info[weird_lead_id][11].lower().strip()

		# Flag that will be true if Type of Lead is empty or is one of the type_of_lead_keywords, which means that we can go in and fix the Specialization 1 and 2 for that entry
		spec_flag = False


		# Now we look into the Type of Lead, and if we find the keyword, we substitute it with the clean type of lead. If Type of Lead is empty, we update the Specializations 1 and 2.
		if(type_of_lead == ""):
			spec_updater(weird_lead_id, spec1, 4, heard, country)
			spec_updater(weird_lead_id, spec2, 5, heard, country)
			continue

		# Only if the Type of Lead has the relevant keyword, we will update it
		for keyword in type_of_lead_keywords.keys():
			if (keyword in type_of_lead):
				spec_flag = True
				# We found the keyword in type of lead that matches the type of lead that was agreed by the marketing team, so we update it accordingly
				dict_tuple_updater(all_weird_leads_info, weird_lead_id, 3, type_of_lead_keywords[keyword])
				break		

		# Now we clean the Specialization 1 and 2 in the similar fashion. We do this here to avoid too much complexity within the loops
		if(spec_flag):
			spec_updater(weird_lead_id, spec1, 4, heard, country)
			spec_updater(weird_lead_id, spec2, 5, heard, country)

		# Else: we simply skip the specialization and type of lead and will let marketing team take care of it later

def export_source_leads():
	# This method exports all the information collected from all sources (XLS or CSV) into external XLS for further import into Access database. The order of data fields that we export the date mataches the schema in the database
	# This method also exports all the weird leads into another XLS

	# IN THIS METHOD YOU WILL CHANGE THE RESPECTIVE COLUMN NAMES IN QUOTATION MARKS!
	global all_csv_entries_info
	global all_xls_entries_info
	global all_weird_leads_info
	global export_column_names
	global export_xls_filename
	global export_leads_sheet_name
	global export_weird_leads_sheet_name

	# Open the file in writing and truncating mode for leads export
	try:
		with open(export_xls_filename, "w+") as all_sources:
			workbook = xlwt.Workbook()

			# Add the relevant sheets
			all_leads_sheet = workbook.add_sheet(export_leads_sheet_name)
			all_weird_leads_sheet = workbook.add_sheet(export_weird_leads_sheet_name)

			# Output CSV and XLS row number
			row_number = 1

			# Weird record row number
			weird_row_number = 1


			# Create the font style - bold - for the top row in the both sheets
			font_style_string = "font: bold on"
			font_style = xlwt.easyxf(font_style_string)

			# Now we need to format the date output
			date_style = xlwt.XFStyle()
			date_style.num_format_str = "YYYY-MM-DD"

			try:
				# Write the column names in bold on All Source Leads Sheet in Row 0
				for name in range(len(export_column_names_list)):
					all_leads_sheet.write(0, name, export_column_names_list[name], style=font_style)
				

				# Write the same columns and add the column for Filename in Weird leads sheet in Row 0
				for name in range(len(export_column_names_list)):
					all_weird_leads_sheet.write(0, name, export_column_names_list[name], style=font_style)


				# # IF YOU CHANGE THE OUTPUT FILE COLUMN NAME, MAKE SURE TO CHANGE IT ABOVE AND HERE AS WELL! 
				
					#all_weird_leads_sheet.write(0, export_column_names["Source"], "Source", style=font_style)
			except KeyError as e:
				print("Leads export error: Unable to find the following column name in export_column_names data structure defined at the beginning of the script:")
				print(e)
				print("Please make sure that the column names are named correctly in export_column_names (lines 199-212). You may refer to the user manual as well.")
				sys.exit(0)				
			

			
			# Begin writing the leads info from CSVs - Unbounce leads. Note that for Unbounce Leads, Specialization 2 field will remain empty at all times

			# For reference...
			# export_column_names = ["Date Added", "First Name", "Last Name", "Email", "Type of Lead", "Specialization 1", "Specialization 2", "How did you hear about us?", 
			# 				"JD/LLB?", "City", "Province", "Country"]


			# IF YOU CHANGE ANY OF THE OUTPUT FILE COLUMN NAMES, MAKE SURE TO CHANGE IT ABOVE AND HERE AS WELL! 
			for csv_id in all_csv_entries_info.keys():
				try:
					# Date
					all_leads_sheet.write(row_number, export_column_names["Date Added"], all_csv_entries_info[csv_id][0], style=date_style)
					# First Name
					all_leads_sheet.write(row_number, export_column_names["First Name"], str(all_csv_entries_info[csv_id][1].split("\t")[0]).title())
					# Last Name
					all_leads_sheet.write(row_number, export_column_names["Last Name"], str(all_csv_entries_info[csv_id][1].split("\t")[1]).title())
					# Email
					all_leads_sheet.write(row_number, export_column_names["Email"], all_csv_entries_info[csv_id][2])
					# Type of Lead
					all_leads_sheet.write(row_number, export_column_names["Type of Lead"], all_csv_entries_info[csv_id][3])
					# Specialization 1
					all_leads_sheet.write(row_number, export_column_names["Specialization 1"], all_csv_entries_info[csv_id][4])
					# Specialization 2
					all_leads_sheet.write(row_number, export_column_names["Specialization 2"], all_csv_entries_info[csv_id][5])
					# How did you hear about us?
					all_leads_sheet.write(row_number, export_column_names["How did you hear about us?"], all_csv_entries_info[csv_id][6])
					# JD/LLB?
					all_leads_sheet.write(row_number, export_column_names["JD/LLB?"], all_csv_entries_info[csv_id][7])
					# JD/LLB?
					all_leads_sheet.write(row_number, export_column_names["Phone"], all_csv_entries_info[csv_id][8])
					# City
					all_leads_sheet.write(row_number, export_column_names["City"], all_csv_entries_info[csv_id][9])
					# Province
					all_leads_sheet.write(row_number, export_column_names["Prov"], all_csv_entries_info[csv_id][10])
					# Country
					all_leads_sheet.write(row_number, export_column_names["Country"], all_csv_entries_info[csv_id][11])

					row_number +=1
				except KeyError as e:
					print("CSV leads export error: Unable to find the following column name in export_column_names data structure defined at the beginning of the script:")
					print(e)
					print("Please make sure that the column names are named correctly in export_column_names (lines 199-212) AND in the code block of lines 1085-1109. You may refer to the user manual as well.")
					sys.exit(0)

			# Now begin writing the XLS entries

			# IF YOU CHANGE ANY OF THE OUTPUT FILE COLUMN NAMES, MAKE SURE TO CHANGE IT ABOVE AND HERE AS WELL! 
			for xls_id in all_xls_entries_info.keys():
				try:
				# Date
					all_leads_sheet.write(row_number, export_column_names["Date Added"], all_xls_entries_info[xls_id][0], style=date_style)
					# First Name
					all_leads_sheet.write(row_number, export_column_names["First Name"], str(all_xls_entries_info[xls_id][1].split("\t")[0]).title())
					# Last Name
					all_leads_sheet.write(row_number, export_column_names["Last Name"], str(all_xls_entries_info[xls_id][1].split("\t")[1]).title())
					# Email
					all_leads_sheet.write(row_number, export_column_names["Email"], all_xls_entries_info[xls_id][2])
					# Type of Lead
					all_leads_sheet.write(row_number, export_column_names["Type of Lead"], all_xls_entries_info[xls_id][3])
					# Specialization 1
					all_leads_sheet.write(row_number, export_column_names["Specialization 1"], all_xls_entries_info[xls_id][4])
					# Specialization 2
					all_leads_sheet.write(row_number, export_column_names["Specialization 2"], all_xls_entries_info[xls_id][5])
					# How did you hear about us?
					all_leads_sheet.write(row_number, export_column_names["How did you hear about us?"], all_xls_entries_info[xls_id][6])
					# JD/LLB?
					all_leads_sheet.write(row_number, export_column_names["JD/LLB?"], all_xls_entries_info[xls_id][7])
					# JD/LLB?
					all_leads_sheet.write(row_number, export_column_names["Phone"], all_xls_entries_info[xls_id][8])
					# City
					all_leads_sheet.write(row_number, export_column_names["City"], all_xls_entries_info[xls_id][9])
					# Province
					all_leads_sheet.write(row_number, export_column_names["Prov"], all_xls_entries_info[xls_id][10])
					# Country
					all_leads_sheet.write(row_number, export_column_names["Country"], all_xls_entries_info[xls_id][11])

					row_number +=1
				except KeyError as e:
					print("XLS leads export error: Unable to find the following column name in export_column_names data structure defined at the beginning of the script:")
					print(e)
					print("Please make sure that the column names are named correctly in export_column_names (lines 199-212) AND in the code block of lines 1124-1148. You may refer to the user manual as well.")
					sys.exit(0)


			# Now extract the data from weird leads data structure and print it to the file
			# IF YOU CHANGE ANY OF THE OUTPUT FILE COLUMN NAMES, MAKE SURE TO CHANGE IT ABOVE AND HERE AS WELL! 
			for weird_id in all_weird_leads_info.keys():

				try:
					# Date
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Date Added"], all_weird_leads_info[weird_id][0], style=date_style)
					# First Name
					all_weird_leads_sheet.write(weird_row_number, export_column_names["First Name"], str(all_weird_leads_info[weird_id][1].split("\t")[0]).title())
					# Last Name
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Last Name"], str(all_weird_leads_info[weird_id][1].split("\t")[1]).title())
					# Email
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Email"], all_weird_leads_info[weird_id][2])
					# Type of Lead
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Type of Lead"], all_weird_leads_info[weird_id][3])
					# Specialization 1
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Specialization 1"], all_weird_leads_info[weird_id][4])
					# Specialization 2
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Specialization 2"], all_weird_leads_info[weird_id][5])
					# How did you hear about us?
					all_weird_leads_sheet.write(weird_row_number, export_column_names["How did you hear about us?"], all_weird_leads_info[weird_id][6])
					# JD/LLB?
					all_weird_leads_sheet.write(weird_row_number ,export_column_names["JD/LLB?"], all_weird_leads_info[weird_id][7])
					# JD/LLB?
					all_weird_leads_sheet.write(weird_row_number ,export_column_names["Phone"], all_weird_leads_info[weird_id][8])
					# City
					all_weird_leads_sheet.write(weird_row_number, export_column_names["City"], all_weird_leads_info[weird_id][9])
					# Province
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Prov"], all_weird_leads_info[weird_id][10])
					# Country
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Country"], all_weird_leads_info[weird_id][11])
					# Source File name
					all_weird_leads_sheet.write(weird_row_number, export_column_names["Source"], all_weird_leads_info[weird_id][12])

					weird_row_number +=1
				except KeyError as e:
					print("Weird leads export error: Unable to find the following column name in export_column_names data structure defined at the beginning of the script:")
					print(e)
					print("Please make sure that the column names are named correctly in export_column_names (lines 199-212) AND in the code block of lines 1164-1190. You may refer to the user manual as well.")
					sys.exit(0)
				except Exception as e:
					print("Weird Leads export error: There has been an issue with exporting the following row of informaiton. Please make sure your data is of correct format and try again.")
					print(all_weird_leads_info[weird_id])
					print(type(all_weird_leads_info[weird_id][2]))
					print(e)
					weird_row_number +=1	
					continue
					

			workbook.save(all_sources.name)
		all_sources.close()
	# Since we want to overwrite the export file, we need to make sure it is closed before we overwrite it again.
	except IOError:
		print(export_xls_filename + " is currently open; unable to run the script while this file is open. Please close this file and re-run the script.")
		sys.exit(0)

def main():

	# The main method of the script. Contains the relevant messages for the user. The method scans through the /Sources directory, looks for the CSV and XLS files and reads them one by one.
	# Once all the files are read, CSV data is cleaned and all the data is exported. Relevant stats are displayed
	global csv_entry_id
	global all_csv_entries_info
	global xls_entry_id
	global xls_unchecked_count
	global all_xls_entries_info
	global all_weird_leads_id
	global all_weird_leads_info
	global osgoode_leads
	global osgoode_leads_emails

	global all_duplicates_id
	global all_duplicates
	global csv_total_duplicate_count
	global xls_total_duplicate_count

	global xls_total_entries_count
	global csv_total_entries_count
	global csv_weird_leads_total_count
	global xls_weird_leads_total_count

	csv_file_count = 0
	xls_file_count = 0
	other_file_count = 0

	#time.sleep(5)
	print("-------------------------------------------------------------------------------------\n")
	print("SCRIPT FOR EXTRACTING THE RELEVANT INFOMRATION FROM UNBOUNCE AND MACHFORMS SOURCES. THE EXTRACTED INFORMATION WILL THEN HAVE TO BE MANUALLY IMPORTED INTO MS ACCESS\n")
	#time.sleep(5)
	print("-------------------------------------------------------------------------------------\n")
	print("MAKE SURE TO HAVE ALL THE SOURCE FILES IN .CSV OR .XLS FORMATS IN /Sources DIRECTORY.\nNote: the records that contain the following emails OR names will be output as \"weird\":")
	#time.sleep(2)
	print("The emails: ")
	print(osgoode_leads_emails)
	#time.sleep(2)
	print("The names: ")
	print(osgoode_leads)
	#time.sleep(5)
	print("-------------------------------------------------------------------------------------\n")
	print("BEGIN EXECUTING SCRIPT. PLEASE DO NOT TURN OFF YOUR COMPUTER\n")
	#time.sleep(1)

	
	start_time=time.time()
	# Scan through the Sources directory and extract the relevant leads information from either XLS or CSV files into local data structures. Catch any exceptions if needed.
	# TO DO: Split up the script into sub classes, clean up the global imports into arrays of strings to be imported
	try:
		with os.scandir("Sources") as sourcesdirectory:

			print("Reading the files in /Sources directory...\n")
			for sourcefile in sourcesdirectory:

				if sourcefile.name.endswith(".csv"):
					csv_file_count +=1
					csv_opener(sourcefile)
				elif(sourcefile.name.endswith(".xls")):
					xls_file_count +=1
					xls_opener(sourcefile)
				else:
					other_file_count +=1	
			sourcesdirectory.close()
	
	except FileNotFoundError:
		print("Unable to locate the /Sources directory. Please ensure the /Sources directory exists and this script is in the same directory and try again.")
		sys.exit(0)		

	if (csv_file_count < 1) and (xls_file_count < 1):
		print("CSV or XLS files were not found in /Sources directory. Please ensure the relevant sources are present in the /Sources directory and try again.")
		sys.exit(0)

	if (csv_file_count < 1) and (xls_file_count < 1) and (other_file_count < 1):
		print("/Sources directory does not have any files. Please ensure the relevant sources are present in the /Sources directory and try again.")
		sys.exit(0)

	print("Cleaning the CSV data for Specialization 1 names...")

	csv_data_cleaner()

	print("Cleaning the final Weird leads list for Specializations and Type of Lead...")
	print("NOTE: MACHFORMS WEIRD LEADS WILL NOT BE CLEANED!")

	weird_data_cleaner()

	print("Exporting the relevant information into All_source_leads.xls...")

	# Export all leads from all sources into one XLS
	export_source_leads()

	# Get the stats
	total_records = csv_total_entries_count + xls_total_entries_count
	total_CSV_normal_records = len(all_csv_entries_info)
	total_XLS_normal_records = len(all_xls_entries_info)
	total_normal_extracted = total_CSV_normal_records + total_XLS_normal_records
	total_weird_records = len(all_weird_leads_info)
	total_duplicates_found = len(all_duplicates)
	


	print("-------------------------------------------------------------------------------------")
	print("Done executing script. The source leads were extracted into All_source_leads.xls.\n")
	print("All the \"Weird\" Leads may be found at \"All weird leads\" sheet in All_source_leads.xls")
	print("-------------------------------------------------------------------------------------")
	print("TOTALS:")
	print(str(total_records) + " entries were processed in total.")
	print(str(total_normal_extracted) + " normal leads were extracted.")
	print(str(total_weird_records) + " weird leads were extracted.")
	print(str(total_duplicates_found) + " duplicates were omitted.\n")
	print("Unbounce/CSV statistics: ")
	print(" - " + str(csv_total_entries_count) + " total Unbounce/CSV entries were processed.")
	print(" - " + str(total_CSV_normal_records) + "/" + str(csv_total_entries_count) + " normal Unbounce/CSV leads were extracted.")
	print(" - " + str(csv_total_duplicate_count) + "/" + str(csv_total_entries_count) + " entries were omitted as duplicates.")
	print(" - " + str(csv_weird_leads_total_count) + "/" + str(csv_total_entries_count) + " weird Unbounce/CSV leads were extracted. ")
	print("MachForms/XLS statistics: ")
	print(" - " + str(xls_total_entries_count) + " total MachForms/XLS entries were processed.")
	print(" - " + str(xls_unchecked_count) + " total MachForms/XLS leads chose to not to receive updates.")
	print(" - " + str(total_XLS_normal_records) + "/" + str(xls_total_entries_count) + " normal MachForms/XLS leads were extracted.")
	print(" - " + str(xls_total_duplicate_count) + "/" + str(xls_total_entries_count) + " entries were omitted as duplicates.")
	print(" - " + str(xls_weird_leads_total_count) + "/" + str(xls_total_entries_count) + " weird MachForms/XLS leads were extracted. ")
	print("-------------------------------------------------------------------------------------")

	# print("----------------------------------------------")
	# pprint(all_duplicates)
	# print(csv_total_entries_count)
	# print(csv_total_duplicate_count)
	# print("\n")
	# print(xls_total_entries_count)
	# print(xls_total_duplicate_count)
	# print("\n")
	# print(all_weird_leads_id)
	# print(all_duplicates_id)
	# print("\n")

	# print(len(all_duplicates))
	# print(len(all_csv_entries_info))
	# print(len(all_xls_entries_info))
	# print(total_weird_records)
	# print("PRINTING CLEAN CSV FILES INFO")
	# print(csv_entry_id)
	# pprint(all_csv_entries_info)
	# print("DONE PRINTING CSV FILE")
	# print("---------------------------")
	# print("PRINTING CLEAN XLS FILES INFO")
	# print(xls_entry_id)
	# pprint(all_xls_entries_info)
	# print("DONE PRINTING XLS FILE")
	# print("---------------------------")
	# print("PRINTING CLEAN WEIRD INFO")
	# print(all_weird_leads_id)
	# pprint(all_weird_leads_info)
	# print("DONE PRINTING WEIRD FILE")
	# print("---------------------------")
	# pprint(admissions_database)
	print("--- The program took %s seconds to execute ---" % (time.time() - start_time))
if __name__ == '__main__':
	main()