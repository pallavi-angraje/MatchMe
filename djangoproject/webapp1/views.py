from django.shortcuts import render

# Create your views here.

from webapp1.next_process import next_p
from django.http import HttpResponse


def index(request):
    return HttpResponse('<H2> WELCOME </H2>')


def start_app(request):
    html = """
    <a href = 'next/'>start</a>
    """
    return HttpResponse(html)

def next(request):
    f = open("ans.txt", "w")
    f.close()
    return HttpResponse()