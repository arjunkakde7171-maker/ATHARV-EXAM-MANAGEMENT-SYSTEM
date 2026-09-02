ATHARVKART V3.8 - SELLER ORDER NOTIFICATION UPDATE

FEATURE:
When a customer places an order, every seller whose product is included in that order receives an in-app notification immediately after the order is successfully saved.

ADDED:
1. Seller notification: "New Order Received"
2. Notification contains order number, customer name, seller's item summary, quantity and seller item total.
3. Each seller receives only notification for products belonging to that seller (multi-seller orders supported).
4. Seller email notification is also attempted when SMTP settings are configured.
5. Notifications menu shows unread count badge.
6. Seller menu includes "My Orders" page.
7. /seller/orders displays seller-specific orders and customer/product details.
8. Customer still receives normal order confirmation notification and email.

RUN:
Use RUN_ATHARVKART_V3_7_FINAL.bat

Default URL:
http://127.0.0.1:5000

NOTE:
Email requires SMTP settings in Admin > Settings. In-app seller notification works without SMTP.
