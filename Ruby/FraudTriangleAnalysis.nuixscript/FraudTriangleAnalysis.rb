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

require 'csv'

dialog = TabbedCustomDialog.new("Fraud Triangle Analysis")

case_date_range = $current_case.getStatistics.getCaseDateRange

# Define the word lists used for each category
opportunity_word_list = "Fraud Triangle Analysis - Opportunity"
rationalization_word_list = "Fraud Triangle Analysis - Rationalization"
pressure_word_list = "Fraud Triangle Analysis - Pressure"

# Make sure these word lists exist
def word_list_exists(name)
	begin
		$current_case.count("word-list:\"#{name}\"")
		return true
	rescue Exception => exc
		return false
	end
end

missing_word_lists = []
missing_word_lists << opportunity_word_list if !word_list_exists(opportunity_word_list)
missing_word_lists << rationalization_word_list if !word_list_exists(rationalization_word_list)
missing_word_lists << pressure_word_list if !word_list_exists(pressure_word_list)
if missing_word_lists.size > 0
	message = "Could not locate required word lists:\n\n"
	message += "#{missing_word_lists.join("\n")}\n\n"
	message += "Please make sure these exist before running this script."
	CommonDialogs.showError(message,"Fraud Triangle Analysis - Missing Word Lists")
	exit 1
end

# Get case earliest and latest date values.  Nuix returns LocalDate objects, Nx date picker accepts
# null / java.util.Date or YYYYMMDD string, so we are going to coerce to string in right format
case_earliest_date = case_date_range.getEarliest.toString.gsub("-","")
case_latest_date = case_date_range.getLatest.toString.gsub("-","")

main_tab = dialog.addTab("main_tab","Main")
main_tab.appendSaveFileChooser("output_csv","Output CSV","Comma Separated Values (*.csv)","csv")
main_tab.appendDatePicker("start_date","Start Date",case_earliest_date)
main_tab.appendDatePicker("end_date","End Date",case_latest_date)

emails_tab = dialog.addTab("emails_tab","Emails")
emails_tab.appendCheckBoxes("search_from","Search From",true,"search_to","Search To",true)
emails_tab.appendCheckBoxes("search_cc","Search CC",true,"search_bcc","Search BCC",true)
emails_tab.appendHeader("Email Addresses")
emails_tab.appendStringList("email_addresses")

queries_tab = dialog.addTab("queries_tab","Queries")
headers = ["Name","Query"]
records = []
queries_tab.appendDynamicTable("named_queries","Named Queries",headers,records) do |record,col_index,set_value,value|
	if set_value
		case col_index
		when 0
			record[:name] = value
		when 1
			record[:query] = value
		end
	else
		case col_index
		when 0
			next record[:name]
		when 1
			next record[:query]
		end
	end
end
dynamic_table = queries_tab.getControl("named_queries")
dynamic_table.getModel.setColumnEditable(0)
dynamic_table.getModel.setColumnEditable(1)
dynamic_table.setUserCanAddRecords(true) do
	next {
		:name => "",
		:query => "",
	}
end

annotation_tab = dialog.addTab("annotation_tab","Annotations")
annotation_tab.appendCheckableTextField("tag_items",false,"tag_template","Fraud Triangle|{category}|{id}","Tag Items with")
annotation_tab.appendHeader("The placeholder {category} will be replaced with relevant category at run-time.")
annotation_tab.appendHeader("The placeholder {id} will be replaced with relevant email address for items found by an email address and 'Name' for named queries.")

dialog.validateBeforeClosing do |values|
	if values["email_addresses"].size < 1 && values["named_queries"].size < 1
		CommonDialogs.showWarning("Please provide at least 1 email address or 1 query.")
		next false
	end

	if values["email_addresses"].size > 0
		values["email_addresses"].each_with_index do |email_address,email_address_index|
			issue_found = false
			issue_message = ""
			if email_address.strip.empty?
				issue_found = true
				issue_message = "Email address #{email_address_index+1} is blank, but cannot be blank."
				break
			end
		end

		if issue_found
			CommonDialogs.showWarning(issue_message)
			next false
		end
	end

	if values["named_queries"].size > 0
		issue_found = false
		issue_message = ""
		values["named_queries"].each_with_index do |named_query,named_query_index|
			name = named_query[:name]
			query = named_query[:query]
			if name.strip.empty?
				issue_message = "Named query #{named_query_index+1} has no name."
				issue_found = true
				break
			end
			if query.strip.empty?
				issue_message = "Named query #{named_query_index+1} has no query."
				issue_found = true
				break
			else
				begin
					$current_case.search(query,{"limit"=>0})
				rescue Exception => exc
					issue_message = "Named query #{named_query_index+1} has an invalid query: #{exc.message}"
					issue_found = true
					break
				end
			end
		end

		if issue_found
			CommonDialogs.showWarning(issue_message)
			next false
		end
	end

	if values["output_csv"].strip.empty?
		CommonDialogs.showWarning("Please provide an output CSV file path.")
		next false
	end

	if values["tag_items"] && values["tag_template"].strip.empty?
		CommonDialogs.showWarning("Please provide a non-empy tag template.")
		next false
	end

	if !values["search_from"] && !values["search_to"] && !values["search_cc"] && !values["search_bcc"]
		CommonDialogs.showWarning("Please select at least one address field to search against.")
		next false
	end

	next true
end

dialog.display
if dialog.getDialogResult == true
	values = dialog.toMap

	output_csv = values["output_csv"]

	# Coerce date picker values to strings which can be used in Nuix date range query
	start_date = org.joda.time.DateTime.new(values["start_date"]).toString("YYYYMMdd")
	end_date = org.joda.time.DateTime.new(values["end_date"]).toString("YYYYMMdd")

	# Lets define this up front since we will be using it repeatedly
	comm_date_range_query = "(comm-date:[#{start_date} TO #{end_date}])"

	# List of email addresses user provided
	email_addresses = values["email_addresses"]

	tag_items = values["tag_items"]
	tag_template = values["tag_template"]

	search_from = values["search_from"]
	search_to = values["search_to"]
	search_cc = values["search_cc"]
	search_bcc = values["search_bcc"]

	named_queries = values["named_queries"]

	annotater = $utilities.getBulkAnnotater

	ProgressDialog.forBlock do |pd|
		pd.logMessage("Date Range: #{start_date} - #{end_date}")

		CSV.open(output_csv,"w:utf-8") do |csv|
			csv << [
				"Query Name",
				"Email Address",
				"Overall Item Count",
				"Opportunity Count",
				"Rationalization Count",
				"Pressure Count",
				"Opportunity Percentage",
				"Rationalization Percentage",
				"Pressure Percentage",
			]

			email_addresses.each_with_index do |email_address,email_address_index|

				pd.setMainStatusAndLogIt("Processing: #{email_address}")
				pd.setMainProgress(email_address_index+1,email_addresses.size)

				# Email address is used to generate a query for that email address in the selected fields
				field_sub_queries = []
				field_sub_queries << "from:(#{email_address})" if search_from
				field_sub_queries << "to:(#{email_address})" if search_to
				field_sub_queries << "cc:(#{email_address})" if search_cc
				field_sub_queries << "bcc:(#{email_address})" if search_bcc

				email_address_query = "(#{field_sub_queries.join(" OR ")})"
				overall_query = "kind:email AND #{comm_date_range_query} AND #{email_address_query}"
				overall_count = $current_case.count(overall_query).to_f

				# Get counts for items which meet our address criteria and have some of the category words in
				# defined in the category terms word lists
				opportunity_query = "#{overall_query} AND word-list:\"#{opportunity_word_list}\""
				rationalization_query = "#{overall_query} AND word-list:\"#{rationalization_word_list}\""
				pressure_query = "#{overall_query} AND word-list:\"#{pressure_word_list}\""

				opportunity_items = $current_case.searchUnsorted(opportunity_query)
				rationalization_items = $current_case.searchUnsorted(rationalization_query)
				pressure_items = $current_case.searchUnsorted(pressure_query)

				opportunity_count = opportunity_items.size.to_f
				rationalization_count = rationalization_items.size.to_f
				pressure_count = pressure_items.size.to_f

				# Do some math to calculate percentages
				opportunity_rating = ((opportunity_count / overall_count) * 100.0).round(2)
				rationalization_rating = ((rationalization_count / overall_count) * 100.0).round(2)
				pressure_rating = ((pressure_count / overall_count) * 100.0).round(2)

				pd.logMessage("#{email_address} - Overall: #{overall_count.to_i}, Opportunity: #{opportunity_rating}%, Rationalization: #{rationalization_rating}%, Pressure: #{pressure_rating}%")

				csv << [
					"",
					email_address,
					overall_count,
					opportunity_count,
					rationalization_count,
					pressure_count,
					"#{opportunity_rating} %",
					"#{rationalization_rating} %",
					"#{pressure_rating} %",
				]

				if tag_items
					resolved_tag_template = tag_template.gsub(/\{category\}/,"Opportunity")
					resolved_tag_template = resolved_tag_template.gsub(/\{id\}/,email_address)
					pd.setSubStatusAndLogIt("Tagging #{opportunity_items.size} opportunity items with: #{resolved_tag_template}")
					annotater.addTag(resolved_tag_template,opportunity_items) do |info|
						pd.setSubProgress(info.stageCount,opportunity_items.size)
					end

					resolved_tag_template = tag_template.gsub(/\{category\}/,"Rationalization")
					resolved_tag_template = resolved_tag_template.gsub(/\{id\}/,email_address)
					pd.setSubStatusAndLogIt("Tagging #{rationalization_items.size} rationalization items with: #{resolved_tag_template}")
					annotater.addTag(resolved_tag_template,rationalization_items) do |info|
						pd.setSubProgress(info.stageCount,rationalization_items.size)
					end

					resolved_tag_template = tag_template.gsub(/\{category\}/,"Pressure")
					resolved_tag_template = resolved_tag_template.gsub(/\{id\}/,email_address)
					pd.setSubStatusAndLogIt("Tagging #{pressure_items.size} pressure items with: #{resolved_tag_template}")
					annotater.addTag(resolved_tag_template,pressure_items) do |info|
						pd.setSubProgress(info.stageCount,pressure_items.size)
					end
				end
			end

			named_queries.each do |named_query,named_query_index|
				name = named_query[:name]
				query = named_query[:query]
				overall_count = $current_case.count(query).to_f

				# Get counts for items which meet our criteria and have some of the category words in
				# defined in the category terms word lists
				opportunity_query = "#{query} AND word-list:\"#{opportunity_word_list}\""
				rationalization_query = "#{query} AND word-list:\"#{rationalization_word_list}\""
				pressure_query = "#{query} AND word-list:\"#{pressure_word_list}\""

				opportunity_items = $current_case.searchUnsorted(opportunity_query)
				rationalization_items = $current_case.searchUnsorted(rationalization_query)
				pressure_items = $current_case.searchUnsorted(pressure_query)

				opportunity_count = opportunity_items.size.to_f
				rationalization_count = rationalization_items.size.to_f
				pressure_count = pressure_items.size.to_f

				# Do some math to calculate percentages
				opportunity_rating = ((opportunity_count / overall_count) * 100.0).round(2)
				rationalization_rating = ((rationalization_count / overall_count) * 100.0).round(2)
				pressure_rating = ((pressure_count / overall_count) * 100.0).round(2)

				pd.logMessage("#{name}/#{query} - Overall: #{overall_count.to_i}, Opportunity: #{opportunity_rating}%, Rationalization: #{rationalization_rating}%, Pressure: #{pressure_rating}%")

				csv << [
					name,
					"",
					overall_count,
					opportunity_count,
					rationalization_count,
					pressure_count,
					"#{opportunity_rating} %",
					"#{rationalization_rating} %",
					"#{pressure_rating} %",
				]

				if tag_items
					resolved_tag_template = tag_template.gsub(/\{category\}/,"Opportunity")
					resolved_tag_template = resolved_tag_template.gsub(/\{id\}/,name)
					pd.setSubStatusAndLogIt("Tagging #{opportunity_items.size} opportunity items with: #{resolved_tag_template}")
					annotater.addTag(resolved_tag_template,opportunity_items) do |info|
						pd.setSubProgress(info.stageCount,opportunity_items.size)
					end

					resolved_tag_template = tag_template.gsub(/\{category\}/,"Rationalization")
					resolved_tag_template = resolved_tag_template.gsub(/\{id\}/,name)
					pd.setSubStatusAndLogIt("Tagging #{rationalization_items.size} rationalization items with: #{resolved_tag_template}")
					annotater.addTag(resolved_tag_template,rationalization_items) do |info|
						pd.setSubProgress(info.stageCount,rationalization_items.size)
					end

					resolved_tag_template = tag_template.gsub(/\{category\}/,"Pressure")
					resolved_tag_template = resolved_tag_template.gsub(/\{id\}/,name)
					pd.setSubStatusAndLogIt("Tagging #{pressure_items.size} pressure items with: #{resolved_tag_template}")
					annotater.addTag(resolved_tag_template,pressure_items) do |info|
						pd.setSubProgress(info.stageCount,pressure_items.size)
					end
				end
			end
		end

		pd.setCompleted
	end
end