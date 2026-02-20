import csv
import win32com.client as win32

input_file = "example.csv"

topics = {}
topics_and_addresses = []

with open(input_file, mode="r", encoding="utf-8-sig", newline="") as file:
	reader = csv.DictReader(file, delimiter=";")
	for row in reader:
		topic = row.get("topic")
		emails = row.get("emails")
		topics.update({"topic": topic, "emails": emails})
		topics_and_addresses.append(topics)

outlook = win32.Dispatch('Outlook.Application')

for n in topics_and_addresses:
	mailItem = outlook.CreateItem(0)
	mailItem.Subject = f"Email about {n['topic']}"
	mailItem.BodyFormat = 1

	mailItem.HTMLBody = f"""
		<html>
			<body>
				<p>paragraph 1</p>
	
				<p><b>{n['topic']}</b></p>
	
				<p>paragraph 2</p>
	
				<p>paragraph 3<b>bolded text</b></p>
	
				<p>signature<br>
				name<br>
				title<br>
				organization</p>
			</body>
		</html>
	"""

	mailItem.To = n['emails']
	# save before sending
	mailItem.Save()

