"""crawler URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/2.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.conf.urls import url
from django.contrib.auth.views import login,logout
from . import views


urlpatterns = [
    url(r'^$', login,
        {'template_name': 'login.html'},
        name='login'),
    url(r'^login/$', login,
        {'template_name': 'login.html'},
        name='login'),
    url(r'^logout/$', logout, name='logout'),
    url(r'^top/$', views.top, name='top.html'),
    url(r'^helpQA/$', views.helpQA),
    url(r'^csv_upload/$', views.csv_upload),
    url(r'^csv_delete/$', views.csv_delete),
]
