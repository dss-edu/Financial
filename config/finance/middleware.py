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

        ip_addr = request.META.get('REMOTE_ADDR')
        username = request.user.username
        school = view_kwargs['school']
        date = now()
        link = request.build_absolute_uri()

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
