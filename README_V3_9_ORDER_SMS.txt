ATHARVKART V3.9 - CUSTOMER ORDER CONFIRMATION SMS

When customer successfully places an order, the system can automatically send an SMS confirmation to the mobile number entered at checkout.

Provider: MSG91 Flow API
Endpoint: https://control.msg91.com/api/v5/flow

Admin > Settings > Order Confirmation SMS:
- Enable SMS
- MSG91 Auth Key
- Sender ID
- Flow ID (DLT mapped)
- Country code (91 for India)

The software stores SMS attempts in Admin > SMS Logs. SMS is triggered only after the order transaction is committed.

IMPORTANT FOR INDIA:
Use an approved DLT sender/template/flow and the correct variables. MSG91 documentation states Indian SMS requires DLT-compliant template configuration.
