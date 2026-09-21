import http.client
import os
import urllib.parse


def invio(messaggio):
    token = os.getenv("PUSHOVER_TOKEN", "")
    user = os.getenv("PUSHOVER_USER", "")

    if not token or not user:
        print("Pushover notification skipped: credentials are not configured.")
        return

    conn = http.client.HTTPSConnection("api.pushover.net:443")
    conn.request(
        "POST",
        "/1/messages.json",
        urllib.parse.urlencode({
            "token": token,
            "user": user,
            "message": messaggio,
        }),
        {"Content-type": "application/x-www-form-urlencoded"},
    )
    response = conn.getresponse()
    print(response.status, response.read())
