app_name = "graph_mail"
app_title = "Graph Mail"
app_publisher = "Ramanathan "
app_description = "Microsoft Graph API mailer"
app_email = "ramanathan.saminathan@infiligence.com"
app_license = "mit"

override_doctype_class = {
    "Email Queue": "graph_mail.overrides.email_queue.GraphEmailQueue"
}
