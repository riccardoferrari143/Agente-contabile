"""
Integrazione Stripe per abbonamenti.
Configura STRIPE_SECRET_KEY nelle variabili d'ambiente.
"""
import os, stripe

stripe.api_key = os.getenv("STRIPE_SECRET_KEY", "")

PIANI = {
    "free":  {"nome": "Free",  "prezzo": 0,    "aziende": 1,  "movimenti_mese": 100},
    "base":  {"nome": "Base",  "prezzo": 1900, "aziende": 3,  "movimenti_mese": 500},   # 19€/mese
    "pro":   {"nome": "Pro",   "prezzo": 4900, "aziende": 20, "movimenti_mese": 9999},  # 49€/mese
}

# ID dei price su Stripe (da configurare nel dashboard Stripe)
STRIPE_PRICE_IDS = {
    "base": os.getenv("STRIPE_PRICE_BASE", "price_BASE_ID_QUI"),
    "pro":  os.getenv("STRIPE_PRICE_PRO",  "price_PRO_ID_QUI"),
}

def crea_checkout_session(email, piano, success_url, cancel_url):
    """Crea una sessione Stripe Checkout per l'abbonamento."""
    if not stripe.api_key:
        return None, "Stripe non configurato."
    try:
        session = stripe.checkout.Session.create(
            payment_method_types=["card"],
            mode="subscription",
            customer_email=email,
            line_items=[{"price": STRIPE_PRICE_IDS[piano], "quantity": 1}],
            success_url=success_url + "?session_id={CHECKOUT_SESSION_ID}",
            cancel_url=cancel_url,
        )
        return session.url, None
    except Exception as e:
        return None, str(e)

def verifica_abbonamento(stripe_customer_id):
    """Ritorna il piano attivo per il customer."""
    if not stripe.api_key or not stripe_customer_id:
        return "free"
    try:
        subs = stripe.Subscription.list(customer=stripe_customer_id, status="active", limit=1)
        if not subs.data:
            return "free"
        price_id = subs.data[0]["items"]["data"][0]["price"]["id"]
        for piano, pid in STRIPE_PRICE_IDS.items():
            if price_id == pid:
                return piano
        return "free"
    except:
        return "free"

def limiti_piano(piano):
    return PIANI.get(piano, PIANI["free"])
