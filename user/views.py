from django.shortcuts import render, redirect
from django.contrib.auth.forms import UserCreationForm
from .forms import *
from .models import *
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.core.mail import send_mail, EmailMessage
from django.contrib import messages
from django.http import HttpResponse
User = get_user_model()
from django.contrib.auth.decorators import login_required
import xlwt


def register(request):
    if request.method == 'POST':
        form = TeamRegistrationForm(request.POST)
        if form.is_valid():
            n = form.cleaned_data.get('team_name')
            teams = Team.objects.filter(team_name=n).first()
            if teams:
                messages.add_message(request, messages.INFO, 'This team name is already taken.')
                return redirect('register')
            
            
            form.save()
            email1 = form.cleaned_data.get('email1')
            team_name = form.cleaned_data.get('team_name')
            form.save()
            messages.add_message(request, messages.INFO, 'Thank You for Registering!  Team leader will receive a confirmation mail.')
            send_mail(
                        'Elevate 21', 
                        f"""Dear { team_name },
I hope this email finds you in the best of your health. We at EDC, thank you for registering for Elevate'21. We are extremely excited and glad that you could join us on this journey to revive the hustle.
With this fiesta around the corner, you will get your regular updates on events and all other further information on the mail itself.
We hope you savor this journey of new challenges, unlimited ideas, and untiring zeal.
Cheers,
Team EDC
                        """, 
                        'esummit@edctiet.com',
                        [email1],
                    )
            return redirect('login')
        else:
            messages.add_message(request, messages.INFO, 'This team name is already taken.')
            return redirect('register')
    else:
        form = TeamRegistrationForm()
    return render(request, 'user/register.html', {'form': form})





# def update_user(request):
#     if request.method == 'POST':
#         form = TeamUpdate(request.POST, instance=request.user)
#         if form.is_valid():
#             form.save()
#             return redirect('dashboard')
#     else:
#         form = TeamUpdate()
#     return render(request, 'user/update-profile.html',{'form': form})


def registerationclosed(request):
    return render(request, 'user/registerationclosed.html')

# def export_answers_xls(request):
#     if request.user.is_superuser:
#         response = HttpResponse(content_type='application/ms-excel')
#         response['Content-Disposition'] = 'attachment; filename="responses.xls"'
    
#         wb = xlwt.Workbook(encoding='utf-8')
#         ws = wb.add_sheet('Elevate Responses') # this will make a sheet named Users Data
    
#         # Sheet header, first row
#         row_num = 0
    
#         font_style = xlwt.XFStyle()
#         font_style.font.bold = True
    
#         columns = ['Team Name', 'Email 1', 'Name 1', 'Contact No 1', 'Discord 1', 'Email 2', 'Name 2', 'Contact No 2', 'Discord 2', 'Email 3', 'Name 3', 'Contact No 3', 'Discord 3', 'Email 4', 'Name 4', 'Contact No 4', 'Discord 4',]
    
#         for col_num in range(len(columns)):
#             ws.write(row_num, col_num, columns[col_num], font_style) # at 0 row 0 column 
    
#         # Sheet body, remaining rows
#         font_style = xlwt.XFStyle()
    
#         rows = Team.objects.all().values_list('team_name', 'email1', 'name1', 'contact_no1', 'discord_id_1', 'email2', 'name2', 'contact_no2', 'discord_id_2', 'email3', 'name3', 'contact_no3', 'discord_id_3', 'email4', 'name4', 'contact_no4', 'discord_id_4', )
#         for row in rows:
#             row_num += 1
#             for col_num in range(len(row)):
#                 ws.write(row_num, col_num, row[col_num], font_style)
    
#         wb.save(response)
    
#         return response        
#     else:
#         return redirect('register')


def export_answers_xls(request):
    if request.user.is_superuser:
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="responses.xls"'
    
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Esummit Responses') # this will make a sheet named Users Data
    
        # Sheet header, first row
        row_num = 0
    
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
    
        columns = ['Team', 'Category A', 'Category B', 'Category C']
    
        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num], font_style) # at 0 row 0 column 
    
        # Sheet body, remaining rows
        font_style = xlwt.XFStyle()

        teams = Team.objects.all()
        for team in teams:

            sellus = SellUs.objects.filter(team=team)

            print(sellus)

            sua = 0
            sub = 0
            suc = 0
            print("###################")
            for sell in sellus:
                print(sell.item)
                if sell.item.category_1:
                    sua = sua + sell.quantity
                if sell.item.category_2:
                    sub = sub + sell.quantity
                if sell.item.category_3:
                    suc = suc + sell.quantity


            sellts = SendRequest.objects.filter(from_team=team).filter(is_accepted=True)

            sta = 0
            stb = 0
            stc = 0

            for sellt in sellts:
                if sellt.item.category_1:
                    sta = sta + sellt.quantity
                if sellt.item.category_2:
                    stb = stb + sellt.quantity
                if sellt.item.category_3:
                    stb = stc + sellt.quantity


            sa = sta + sua
            sb = stb + sub
            sc = stc + suc



            row = [team.team_name, sa, sb, sc]
            row_num += 1
            for col_num in range(len(row)):
                ws.write(row_num, col_num, row[col_num], font_style)
    
        wb.save(response)
    
        return response        
    else:
        return redirect('register')

