script_directory = File.dirname(__FILE__)
require File.join(script_directory,"Nx.jar")
java_import "com.nuix.nx.NuixConnection"
java_import "com.nuix.nx.LookAndFeelHelper"
java_import "com.nuix.nx.dialogs.ChoiceDialog"
java_import "com.nuix.nx.dialogs.TabbedCustomDialog"
java_import "com.nuix.nx.dialogs.CommonDialogs"
java_import "com.nuix.nx.dialogs.ProgressDialog"
java_import "com.nuix.nx.dialogs.ProcessingStatusDialog"
java_import "com.nuix.nx.digest.DigestHelper"
java_import "com.nuix.nx.controls.models.Choice"

LookAndFeelHelper.setWindowsIfMetal
NuixConnection.setUtilities($utilities)
NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

dialog = TabbedCustomDialog.new("Fraud Triangle Analysis")

case_date_range = $current_case.getStatistics.getCaseDateRange

# Define the word lists used for each category
opportunity_word_list = "Fraud Triangle Analysis - Opportunity"
rationalization_word_list = "Fraud Triangle Analysis - Opportunity"
pressure_word_list = "Fraud Triangle Analysis - Opportunity"

# Get case earliest and latest date values.  Nuix returns LocalDate objects, Nx date picker accepts
# null / java.util.Date or YYYYMMDD string, so we are going to coerce to string in right format
case_earliest_date = case_date_range.getEarliest.toString.gsub("-","")
case_latest_date = case_date_range.getLatest.toString.gsub("-","")

main_tab = dialog.addTab("main_tab","Main")
main_tab.appendDatePicker("start_date","Start Date",case_earliest_date)
main_tab.appendDatePicker("end_date","End Date",case_latest_date)
main_tab.appendHeader("Email Addresses")
main_tab.appendStringList("email_addresses",)

dialog.display
if dialog.getDialogResult == true
	values = dialog.toMap

	# Coerce date picker values to strings which can be used in Nuix date range query
	start_date = org.joda.time.DateTime.new(values["start_date"]).toString("YYYYMMdd")
	end_date = org.joda.time.DateTime.new(values["end_date"]).toString("YYYYMMdd")

	# Lets define this up front since we will be using it repeatedly
	comm_date_range_query = "(comm-date:[#{start_date} TO #{end_date}])"

	# List of email addresses user provided
	email_addresses = values["email_addresses"]

	ProgressDialog.forBlock do |pd|
		pd.logMessage("Date Range: #{start_date} - #{end_date}")

		email_addresses.each_with_index do |email_address,email_address_index|

			# Email address is used to generate a query for that email address in any of the communication fields
			email_address_query = "(from:(#{email_address}) OR to:(#{email_address}) OR cc:(#{email_address}) OR bcc:(#{email_address}))"
			overall_query = "kind:email AND #{comm_date_range_query} AND #{email_address_query}"
			overall_count = $current_case.count(overall_query).to_f

			# Get counts for items which meet our address criteria and have some of the category words in
			# defined in the category terms word lists
			opportunity_count_query = "#{overall_query} AND word-list:\"#{opportunity_word_list }\""
			rationalization_count_query = "#{overall_query} AND word-list:\"#{rationalization_word_list }\""
			pressure_count_query = "#{overall_query} AND word-list:\"#{pressure_word_list }\""

			# Do some math to calculate percentages
			opportunity_count = $current_case.count(opportunity_count_query).to_f
			rationalization_count = $current_case.count(rationalization_count_query).to_f
			pressure_count = $current_case.count(pressure_count_query).to_f

			# TODO: Do these make more sense as percentages like 83% rather than 0.83?
			opportunity_rating = opportunity_count / overall_count
			rationalization_rating = rationalization_count / overall_count
			pressure_rating = pressure_count / overall_count

			# TODO: Replace this debugging log message with reporting, annotations, etc...
			pd.logMessage("Overall: #{overall_count.to_i}, Opportunity: #{opportunity_rating}, Rationalization: #{rationalization_rating}, Pressure: #{pressure_rating}")
		end
	end
end