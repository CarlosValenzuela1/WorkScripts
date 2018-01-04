;This script is for all the ticket closure forms

;;============================================================================================;;
;MAIL1 Sends an email to L2 when they DID NOT SENT EMAIL to user
::mail1::
(
TICKETNUMBER - Moved back to In Progress - Please use ticket closure pilot procedure

Hello!  

This is Carlos Valenzuela from the Customer Care team.  I've reviewed this ticket and found that the resolution email communication and user agreement to close is not documented in this ticket.  
As a result, I have changed this ticket status to Status = In Progress.

Can you please refer to page #6 of the Ticket Closure pilot procedure document and follow the procedure to ensure that this user is informed of the resolution and that the ticket is resolved to his/her satisfaction?
http://ithub.cargill.com/serviceoperations/global-infra/quality/Shared%20Documents/Continuous%20Improvement/Impove%20User%20Satisfaction%20with%20Ticket%20Closure/TicketClosure-PilotProcedures-SupportGroups.docx

Thank you!

Kindest Regards/ Saludos,

Carlos Valenzuela

Service Desk EMEA 
TATA Consultancy Services
Cargill
Budapest
Web contact: http://www.helpdesk.cargill.com 
Service Desk feedback:  ITEscalation_EMEA@cargill.com
)
;;============================================================================================;;



;;============================================================================================;;
::rem1::
(
Customer Care action summary - Emailed Assigned Analyst - Requested to follow pilot procedure documentation.
)

::rem2::
(
***********************
Customer Care action summary – audit passed
Carlos Valenzuela
Resolution email communication was sent and user agreement to close is documented in the ticket
***********************
)
;;============================================================================================;;



;;============================================================================================;;
::chat1::
(
Hello!
I am contacting you from the "Ticket Closure Team"
We are trying to improve the customer satisfaction of our users.
Your ticket: ()
Has been set as resolved, do you agree with this decision?
Thank you!
)
;;============================================================================================;;



;;============================================================================================;;
::remwwts::
(
Customer Care - requested agreement from the user, since it's a WWTS ticket
)
;;============================================================================================;;