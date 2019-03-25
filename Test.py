from django.shortcuts import render, redirect
from employee.forms import EmployeeForm
from employee.models import Employee
import pandas as pd
import json
# Create your views here.

from django.http import HttpResponse
from django.http import JsonResponse


def read_data(request):

    released = {
        "iphone": 2007,
        "iphone 3G": 2008,
        "iphone 3GS": 2009,
        "iphone 4": 2010,
        "iphone 4S": 2011,
        "iphone 5": 2012
    }
    return HttpResponse(released, content_type="application/json")
