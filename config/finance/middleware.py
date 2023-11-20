from django.utils.timezone import now
from .connect import connect

class VisitorMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        return response

    def process_view(self, request, view_func, view_args, view_kwargs):
        cnxn = connect()
        cursor = cnxn.cursor()

        username = request.user.username
        date = now()
        ip_addr = request.META.get('REMOTE_ADDR')
        link = request.build_absolute_uri()

        # local testing should be ignored
        if "127.0.0.1" in ip_addr or "localhost" in link:
            return None

        try: 
            school = view_kwargs['school']
        except:
            return None


        query = 'INSERT INTO [dbo].[Access_Logs] (ip_addr, username, school, access_date, link) VALUES (?, ?, ?, ?, ?);'

        try: 
            cursor.execute(query, (ip_addr, username, school, date, link))
            cnxn.commit()
        except Exception as e:
            print(e)
        finally:
            cursor.close()
            cnxn.close()
        
        return None
