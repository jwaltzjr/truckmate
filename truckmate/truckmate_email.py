import env
from krc.krcemail import KrcEmail

class TruckmateEmail(KrcEmail):

    def default_password(self):
        return env.email_password.value