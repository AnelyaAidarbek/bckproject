from django.shortcuts import render,redirect
from .forms import  DataForm
from .models import  Data
import openpyxl
# Create your views here.
def home(request):
    if request.method=="POST":
        form = DataForm(request.POST,request.FILES)
        if form.is_valid():
            form.save()
            return redirect("Home")
    form = DataForm()
    data = Data.objects.all()
    ctx={
        "form":form,
        "data":data,
    }
    return render(request,"excapp/index.html",context=ctx)

def detail(request,id):
    data = Data.objects.get(pk=id)
    wb = openpyxl.load_workbook(data.file)
    worksheets=wb.worksheets[0]
    excel_data = list()
    header_data=[]
    for index,row in enumerate(worksheets.iter_rows()):
        if index==0:
            for head in row:
                header_data.append(str(head.value))
        else:
            row_data = list()
            for cell in row:
                row_data.append(str(cell.value))
            excel_data.append(row_data)
    context={
        "excell":excel_data,
        "file":data.file,
        "headers":header_data,
    }
    return render(request,"excapp/detail.html",context=context)