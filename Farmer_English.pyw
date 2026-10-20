"""Farmer – herd pedigree, relationship, inbreeding and mating plan.

Start: double-click this file (requires Python 3.8+ from python.org).
Data is saved in herd.csv next to the program (it also opens in Excel).
"""
import csv
import json
import os
import shutil
import sys
import traceback
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont

def _data_folder():
    """Where the program keeps the herd, settings and backups.
    - run as .py: next to the file;
    - portable .exe that already has data next to it: there;
    - installed program: Documents\\Farmer English (Windows does not allow writing into the program folder)."""
    if not getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(__file__))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    for name in ("herd.csv", "herd_pedigree_settings.json"):
        if os.path.exists(os.path.join(exe_dir, name)):
            return exe_dir
    docs = ""
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(260)
        if ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf) == 0:  # 5 = Documents
            docs = buf.value
    except Exception:
        pass
    if not docs:
        docs = os.path.join(os.path.expanduser("~"), "Documents")
    folder = os.path.join(docs, "Farmer English")
    try:
        os.makedirs(folder, exist_ok=True)
        return folder
    except OSError:
        return exe_dir


APP_DIR = _data_folder()
PODESAVANJA = os.path.join(APP_DIR, "herd_pedigree_settings.json")
KOLONE = ["id", "markica", "vrsta", "rasa", "pol", "datum_rodjenja", "otac", "majka", "status", "stado",
          "poreklo", "poreklo_od", "stare_markice", "uneto"]
REDOSLEDI = ["Last added first", "First added first", "By tag", "Youngest first"]


def sada():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
POREKLA = ["Home bred", "Purchased", "Imported"]
VRSTE = ["cattle", "sheep", "goats", "pigs", "horses"]
NAZIVI = {   # species prefix (English and Serbian) -> (male, female, young, young plural)
    "cat": ("bull", "cow", "calf", "calves"), "cow": ("bull", "cow", "calf", "calves"),
    "bul": ("bull", "cow", "calf", "calves"), "gov": ("bull", "cow", "calf", "calves"),
    "she": ("ram", "ewe", "lamb", "lambs"), "ewe": ("ram", "ewe", "lamb", "lambs"), "ovc": ("ram", "ewe", "lamb", "lambs"),
    "goa": ("buck", "doe", "kid", "kids"), "koz": ("buck", "doe", "kid", "kids"),
    "pig": ("boar", "sow", "piglet", "piglets"), "sow": ("boar", "sow", "piglet", "piglets"),
    "svi": ("boar", "sow", "piglet", "piglets"),
    "hor": ("stallion", "mare", "foal", "foals"), "kon": ("stallion", "mare", "foal", "foals"),
}


# logo programa (PNG, ugrađen da program radi kao jedan fajl)
LOGO = {
    64: (
        "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAJL0lEQVR42u2bbYxcZRXHf+e5986dO7Mzu33ZbaFLEShtwYZKIZBA"
        "QUhQoVU+SIyBGsEaNSRIQhBj0C9EhA+EBCVEkQQwkRAoUSKiNVGgtNAADS8tUopaKG1puy3dZeflztyX5/hhZpftdgvb3dnubtMn"
        "mWTm3pmbOb97nv/znHPPEY4wVHFESAH6XmZm4HCVKlday7lqOcVCm4BhagwrUBbDTjG8YYS1YcrfOy7i4HBbhg8ZwXABRARbXsdJ"
        "js/NAqs8l1OMgSSGOAVrmVLDGPAccL3Gf4sTdio8ltb5TduX2aOKAVQEPSKApvGIoOUN/MDzuCuTY3ZchXqERVAUIzIyvEkeqgoI"
        "FkX8DMbLQVTlQBxze9tyHhpq32EABk7yAk6Y5cEgz+p6BeKEBHBEppzBn0MDBVLPxfXzEFZ4+JUaP7rsssZUGIAgg8avwdCJhBn+"
        "FHTwjfJBEgXHTDPDDxMHRQXStpm4YR/PBBHfZD/Kt7AiqAwVif4XebQwi+tLHxMbg8dxNKwlLszCK33MH4qXcsOAzUafbLwpr2d1"
        "YQbXVw4ef8Y3RdKrHCQuzOD68npWi5DqkziiipRfptMRthlDMU5gCi1vrdUFsJ4L1tKfKovaLmK/kYay3xIU6Yhi7FQz3qqQ2k9f"
        "VscuSQImirFBkQ5SbhFBpfd5OhyXrX6GOXGMMoUApCoEruK5Q+6ihUrUNGhsLKznIfWIfYWAxa6f4SrfZ241xIq0znhVxny3FDAC"
        "RV/ZdiDHC+93sbs/oOgnXND9MRfNP0hiIUnHBMFEETYfMLdcZYVrLSuNQRm2Qxrn0oPvQMYb3yUffPU0HnhlAaW6hzTB/O61BXz1"
        "jL388oot5DPJ2CAIagwqwkoXYVmSINKinZ1VyHmwq99n3ftdVGMXEWW0eC1CxlF29AU89tYXyHkpM4No8OcC/OXdeUSp8MDVbzSN"
        "16PVAkliBGGZi9IdJw3XaIVgBZ6yeW+RG585nz2lLEbGNn0coxT9GFUhsYdeZG4h5F/b5/LP/3ayYlEPpZrgmKOCYOIUULpdhHyr"
        "A5v7Ni5kXzlLV75OOlbVVo74W6uCiLLugy5WLuoZk+9aCyrk3VapvgKuUUp1w+7+gJyXEFuDKi0fquAAfWEG1bHPXQHjTsTyJYCq"
        "HDY3j8pNP8MDRBrnZuXqiGkAGSuFCQFwJLcthd6o5cqIUvCTwTt+GGDg8tP3NTxgHPI94QBEIEmFjiDituVb8RzFfobbWiDjwI7e"
        "gN9vOoOMY8m6dhCcKuwtZ7nm7J1cfvoBKnXBiE5dAAN3P+8lXLd0J647ylVL4ORiyK9fXsT+qj942HdTVi3dwS8ue4ckPfolcFKn"
        "QF/okHUt6WiES2DV0l1ceup+Nnw4m49KAYVMwnnzDrLspH5qCaR2fO5/TAEMiKBjGpui0fzv/rrQ1Vbn2qW7PwWZQjlqLIPSgq3b"
        "MQVw1MBEiVOoJ3KIpoxnzk8rAAMGO6ITdv3jMvExrT3AqqA64ObSWDJb7PZTFoBVIecpjgNxAiKK23wfJjIhU+GYAkhVmq+RV4Gi"
        "r7y5p8jT78zjf71tGFHO6uznmi/u4sxZVUrj3PRMKgARaPfTRnpLR1ajh187lXs2LKaWOHimEaKue7+LJ7bM584rtrBi0T7KLYYw"
        "4QAGYvu+msdP1i7FMXrY3t4I1BLD89vn4BplZhBhm98pihImDj/9x1LmFTeyZE6JMG4dhGPiAUaUWuzy120nDwYvekjk2DhW8BNQ"
        "DkmAWBWybkpv6PPI66dx38rN03UKKO1+/LkaMfJxQ9ZL+XdPB72hIRjtdnoqiuBYsy2CEqWGKDHkPEurUrjHFIAMDeaHB/fCiLH/"
        "gPek1jA7V6c9mzSCoOk2BRSIU3NEMtY2ihyybko6RAOkmWqrxC5fW7CHIAP94VEnQSc3IRKnhs58jbu+8haZZkJkuGckarjzhbPZ"
        "sreDmUHUSKU3RXBPKcvlp/ew6ksfUq2DMdMsGFKFjGNZdnIfWZfDBEwVHAcevHoTd794Fut3dBJGHqAUMgnfPfcDblu+Dc9Y4pSW"
        "hMHHfgooVCKH1I6s4BpBV77O/V9/k20H8nzQm8eIsmBWmTNmhoTN2iRpcbnGMRVBI4oRZcQkrkCUQj0VzpxV4azOCtAwulRvXQJk"
        "SgdDjedzShh/+qDSoBMWCU7ZhMhEGnwiIXICwAkAJwBMGAAdouYjKTyAHq8AFHCkYWiUmMaSJjq4fg8cc2RqQTCtu5AAwh3PLWFP"
        "OcB3U8LIIYwcfDdlTzngjueWANL87nEEILVC3lce39zNmrfnk/MSKpHLwtklFs4uUYlccl7Cmrfn8/jmbvK+HhLxTTaAcRfIOEap"
        "RsKft55C0Y+pJQ43XfgeT127kaeu3chNF75HmDgU/Zint55CNWpdODvOaWsNSsWMww9UG40K+ysZespZUhXmt1f5/nnbSVIlSZXV"
        "523n1PYqVoV95SwHKhk8hwkpnxn1nTcgSsUg7GpWYtrxQHAGBU+JrRBbwXMacJLmZ5rnjeikGg9YzwGEXQbldddFdYziLAKxhc58"
        "RHexigAflQLufWkxfbUMfbUM9760mI9KAQJ0F6t05iNiy4REd6N0fXU9FOV11xietZZVKGNujbBW8DPKdefs4OZnZ9OZr/Pklvk8"
        "t30OAAcqPvlMyv6Kz3Xn7MB3oV6bRB1QxFpElWdFN9FeCnl3vMXS2iySvGfDQn776gIC1w7dGhEmDjde8B9uu+Q9wmaBw2S5f8ZD"
        "as1iaQEobeDutg5+VuolMTK+EDnw4G/b5vDE2/PZ9UkOgO72Kt9e8iErFu0jjCdX+a2SFGbgVvu4O7+c21veMGEVCj4kFnpDB4AZ"
        "QYproFRnTKWzrVz2DmuYYA2mcDE9NuVWv5GGS8eXzGjU9tRiKGRSCpmUWtw4ZiZ57yNK6ucxccqthYvpYeB+DDQQldbzSNtMbmhV"
        "09TQ53+TPQaapsoHebRwCd8bsPlE29yQzcxg42TV56FcGzccV42TVR4JQn7ISI2TwyEMaZ39VSZH5zRvnd0fx/z8c1tnh0EQEWzP"
        "Ok5qz/JjVb4z7ZqnY3aK4Y+f1Li/a7TN08NADG+fv1ItV1rL+QrzrE6d9nkFa4SywG4xbDKGtWHK2tG0z/8fi/hk+RsT6zEAAAAA"
        "SUVORK5CYII="
    ),
    36: (
        "iVBORw0KGgoAAAANSUhEUgAAACQAAAAkCAYAAADhAJiYAAAFPklEQVR42r2Y22scVRzHP79zZmYvWTdla2vTatW20VoKikoRalC8"
        "gaIoPvgPiNoqgsY38UVBwQf1SfACPog+GBAExQv60FIVMQhWbUmaxoLXGtskbjK72bmcnw8z1tBu0t3u6oFhHuYw53u+399dyJcq"
        "RgQ3+ymbBwZ5wKXc5RzbVCkDQn+WitAwhqPG8n74F6/XbuWnf84GMABjY1gRXPg5eyuDjAc+T3mWKwUG+ggGQAQGPMuVgc9TlUHG"
        "w8/ZK4IbG8MCiCpWhHTxAC8O1Hi8VYc4IQWMSF/BsEwNBZzvYQtVCGd5qTLCqCpWAOoHeOi8tbwSzhI7xTN9AqIqOS3a9rtT1AjJ"
        "QA1/4SR7qiO8KuGXbEL5wVqqcQKSy7jqQR1uKHjZu5Vkwkv7rc73IE2pI+w0LmVvucyaKEY7AeNUsMKKjxEIPCj6cGyuzLG5MkUf"
        "PAPa5iYCJorRcpk1OPZ4AncnMdqJvQhQ9pV6y+C0/XYjMB97PH9gO/uOrQfghktnePqmQxS9hFTPZEoESWJUlXs8YGscI3IWb1IF"
        "z8ILXwzzweQmAuva31igGVv+WCxSDhKsKGPfb2ZDZYknb5yk3hSs0TNYimMAtnoIJadnl6nkKdOzZd74ZguSSyOibW9hRKkWY1KX"
        "fa2VIsZ/qbEUcQaYZQYOQsnrJoC0EkPgKUYcIDQiS6pnaq1A2U8JrMOIMp9Yzh9oEXgQRoKRlRnwuo1qWRARVOGJ3RNcOLhEK5Hc"
        "tTM55pd8Xv5qmOOLJQSlVo548NppnJ79DK+XOHPr1uNs2bAE8bJ4rln837Vplo+nNgBw+2W/s21tg0bEquz0DGgh8omaLRrJaTIo"
        "bK2FPHb9NABxTA7mP2bIiGKN5vEnA5Rq5q+N2KA5IyJnZ6YvgFIVUpc9KpmNVQv5waKkKTTizpjpC6DBQkxQVoLo32j34eR6Ppka"
        "ouCl3LfzZ67e+Bdh1DmocwAkQCbDm99eyrqBiCiFwMLvCwXe+X4ziRMU4cPJjbxx79dcvXGeZiwdydY1IF3mUG8dvITUZVKpZiyc"
        "V8jCvhXlRKPA2wcvZteF8+j/IZlvHN6ydGxzA3ZOEKP4xhFGFtXOq7yeAG1fV6fkpziXBcS5ps/hPwcp+wmxs4Sxx23bjkPOIP+F"
        "20tOvio8e8tBrrigwVKUJd7FyPLiF5ez79h6fON49Lop7tnxG+EqOayvDMXOkCQQOSF2SsGmPHPzYU6Ek3hGWVNyhFF3RXlPgCTP"
        "b5LXQanCQkuoBCkK1FuCFe3qnz0Bag9QSfRfI+92dQXo9Cpxpaqxlw6hK0AlP8WIYgw4ByUv7XuLZDoNhtbA+K816q2A+pJPvRUw"
        "/lsNa0D/T0Cal69HT5Z4bv8Ohtcu8PjIEYbXLvDc/h0cPVmi5OmpHqx3QEpztcTnAN+DwzNVZsIio7sneOSOaUZ3TzATFjk8U8X3"
        "sn29lTKA0jTAtO+jusI/hcxehqpNAuv46MgQ300M8tGRIQLrGKo2ca43Q1Zwvo8C054I73k+O7WFE2lfhDUi4aqhOvdfM80rXw/z"
        "7qGLaKWWPbumuGqoTiOSjguwlXp9z8dEEe9J+BmbKK3eSmtOqWdg34/rmDhRZfv5dW7c8ieJy1oY6YWdZa10V8MGBSpB7goOFqPe"
        "pGo7bOh2HJOqnCqK7DnKtNo4xoiQjo1hKyOMNuZ4WCwzlQq2WEDa2ZQ9VdifGxgRKBaQSgUrlpnGHA9XRhjNh2apnD7Sa+znIlvg"
        "QYU7nWPYuf6O9IyhYQxTAh+kLV4r38DPy0d6fwNJMFYehxw5cQAAAABJRU5ErkJggg=="
    ),
    32: (
        "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAE8klEQVR42q2Xz29UVRTHP+feN29m6EzrFC2glRjAH4EWiIlSEQQx"
        "xMTERF2wYQNRY6KJMbhkYUz0XzCS6E43szFGZaExkaAJVo3EFrQlWKtCgQaw8zrTmfveu8fFtAilrdOBs3lv8d453+8933POPQJQ"
        "LmP37SMtv0X47DO8aGB/mtDvlSIg3JqpESIbMOTh48+O8uG+t3FzMUUVI4K/8hX9HV18EGZ4NE3AxeA9t8WMgTADNgDnGKxWeKl7"
        "L0OqGAGIjrMpzHI8DClFEamAiGC4jaaKV9BiEescV12DncWdnDI6Shb4KAwpTUfERrDtBleFVOXa83oTwRjBTkfEYUhJlY90lGxQ"
        "u8iBQoGtUURqhMyizgGvsqggVCGXUUKrxClkrFJzzf/mAclEEWmxyNbaRQ4EKhxME1SWEJsCgUAubJ7jQsGthbErOd77/n7+nFrB"
        "E/dNcvDhMUT1ZhAgaYKqcDBQpc85RBYhp4AViFzAu8ce4FwlT2BudKoKGQsjkwXOXi2SzyQcG+shToU3tp+l0hCs6A3pcA5U6QtE"
        "6FiI1Zx5LxTySnl4De//sIHuvMOrzEKbe4KqkM+k3LmiAShW4OvfV/PKI2dvCH7Nr4IIHUGrAqu5gM5sTGcuxiWG1N+YNAEaaVO7"
        "2SAlcgE9HXXCAGpOMLIwy5YBzDmo1DPs3/wHz/dNUK0LxuisAOHb8ZUcGdzAVD1kfWma1x8bZanTXRaAOZaJF3q7avT3ThHPgL2u"
        "YLesnuKpdReYiPL0rZqiO58wE7Mo+2UDmAPhEoN3QtUJ1ui1nHoV1pVqPHRXjXoMtZjZVsftA9CsjGaz0eukWAgVa5U4gal6E5hp"
        "YYoE7XS7fJBi80qXKsY2Ufx4rpPfJjvpX/0P/aumqbnWxliwPOYQBp5Tk518ObyKqhPyGWX4Yicf/rSOemLJWs87e3/hhY3nqS6h"
        "/rYAeBVWZFK+GL2HT3/tvb4NkM+kdIQJUSPDkcH17F1/gYz1zXpfqrpuZeh4L6TazHdgFFUhtCn1xBKn0tJFYtll2EgNW9dcZWNP"
        "RCMRslY5c7nAN2M9hIGn6gL2bxlnZUfKdOM2p8CIMuMsT2+Y4MD2v9AZEAu1hvDxybUMX7qDh+++wr6+v/+3/tvvAwIzcUA6I1Tq"
        "TYaBUV5+dBwYB6DaaN1fW33AiGJNc+DY2clYmflvNtgWmC8bgOrC73PamOuIy7WgVfVlA38tp9nA3/pduVUARiBJ4NSlLqZd8/NT"
        "l7pIElpqtbcEIFWhmFWOjvRQHl7LawNnADgyuIEn113kmQcvEc277SxbT6pUF2WiIAZOTpQo5R2H94xweM8Ipbzj5EQJMdx862xZ"
        "yNCMLQyHIaqKX6jkvIdtvZeZrGY59Hk/hz7fzGQ1y7bey3jf/KadHSEMURGG7eGD2GyOZxsOP38fEIE4gfV31ujKOT45fS/j/3Tw"
        "5uMjPLfpPHXXng4UfC6HaTjeER0lG13gRLGTrdMRsSyyG3SEzTkP0JVTqq7tDSkuFMlEFU4WVzMgAPUTbBTDt0utZl6FYLbWEy8t"
        "tdmlVjP17MgNcNqoYnIDnK7W2OUcg8UCNp/DzD9aI0qqkCrLDm4E8jlMsYB1jsFqjV25AU6rYowIvlzGdu9m6NMv2Fl3vJp4vkOo"
        "zNe4tLerK0Il8XxXd7w69DM7unczVC5jRfD/AoiFQTDA2f8+AAAAAElFTkSuQmCC"
    ),
    16: (
        "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAACEUlEQVR42m2Tv2sUURDHP7Pv7d3t5hRjLviz8Uc4EJQUlsFOEQVJ"
        "Kf4P9v4BFrYBS0srLRQrCak0GCFNRFQCRpEIiqhJzrvbe7f73lhsAnfZTDm872fezHdGALZfcT1NeFDkXAgBCwgjoUAIgghqjRY2"
        "5mM/496RK7yU7jLX4pgX1lDvD0qlKsgIIhJIapAX4ApIEyg8Ls+5FamyEBvqvQy/p6nZEmAiiA24IuLp+5N83UqoWej28bGhrsqC"
        "/HtNABBBggppTXm4co7nn04jogjCwEds/Gly+dRfntx+g6qiigJYkfF+I4G336fYGcTcaP/AB6HjLDuDmDOTXYwoedmiAFj2hexC"
        "Lh7b5v7ND+DKV19+NWilriw7WpADIigEFVDIPaxtHmIqHdKIlcKPW1T5gSok1vPu5yR3n83Sc5aVzSkuHd/m0fwqRgJhxOcqoBwo"
        "rojY3EkZ5IZ+Xs4gaOnMaB8VQCTQzw2zJ7Z4fGeVrA/L31qcPdoliQPDYnxH7EEzsLLrUQ5FEK6e/83QUxFXAAJ4hY6LAfChzHWc"
        "EKEVMUCkWu5FUGGirix9btGe7tCe7rC00aJZV4Rq5T1dBKxPJIiq+qGHmVaPLDdkuWGm1cMVIPvcD4qfSBBgfeyYehk0LHScAeBw"
        "3ZcAGXcpbYwcU3OOxcwxH2DNGvJBgaaxJ42rYkBNRB5gLXPMN+dY/A/S4PZF4nAFhQAAAABJRU5ErkJggg=="
    ),
}

# boje
BG, SURF, SURF2 = "#F3F5F1", "#FFFFFF", "#E9EDE6"
INK, INK2, LINE = "#18221C", "#56645A", "#D6DDD3"
TAG, TAG_INK, BRAND, BRAND_H = "#F4C400", "#1B1A12", "#2E5E45", "#24503A"
OK, OK_BG, WARN, WARN_BG, BAD, BAD_BG = "#23804A", "#DDF1E3", "#9A5C00", "#FBEBCB", "#B3322B", "#F8DEDA"


def nazivi(vrsta):
    return NAZIVI.get((vrsta or "")[:3], ("male", "female", "young", "young"))


def norm(t):
    return "".join(c for c in str(t or "").upper() if c.isalnum())


def kolona(k):
    k = (k or "").strip().lower()
    for a, b in (("đ", "dj"), ("č", "c"), ("ć", "c"), ("š", "s"), ("ž", "z")):
        k = k.replace(a, b)
    return "_".join(k.split())


def stare_lista(r):
    """Ranije usne markice grla (kolona stare_markice, razdvojene zarezom)."""
    return [x.strip() for x in str(r.get("stare_markice", "") or "").split(",") if x.strip()]


def sort_key(s):
    """K2 < K10 (prirodno sortiranje)."""
    import re
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", str(s))]


def parsiraj_datum(t):
    t = str(t or "").strip().rstrip(".")
    if not t:
        return None
    try:
        t = t.split(" ")[0]
        if "-" in t:
            g, m, d = t.split("-")[:3]
            return date(int(g), int(m), int(d))
        for s in (".", "/"):
            if s in t:
                d, m, g = t.split(s)[:3]
                return date(int(g), int(m), int(d))
        if len(t) == 4 and t.isdigit():
            return date(int(t), 7, 1)
    except (ValueError, TypeError):
        pass
    return None


def fmt_datum(d):
    return d.strftime("%d/%m/%Y") if d else ""


def starost_meseci(d):
    if not d:
        return None
    n = date.today()
    return (n.year - d.year) * 12 + n.month - d.month - (n.day < d.day)


def fmt_starost(m):
    if m is None:
        return ""
    if m < 24:
        return f"{m} mo."
    g = m // 12
    return f"{g} " + ("year" if g == 1 else "years")


def fmt_upis(t):
    """'2026-10-05 14:30:00' -> '05.10.2026.'"""
    try:
        return datetime.strptime(t[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return "–"


def pct(f):
    return (f"{f * 100:.2f}" if 0 < f < 0.1 else f"{f * 100:.1f}") + "%"


# =============================================================== podaci (CSV)

class Tablica:
    def __init__(self, putanja, ucitaj=True):
        self.putanja = putanja
        self.redovi = []          # lista dict-ova sa kljucevima iz KOLONE (+ dodatne kolone)
        self.dodatne = []
        self.mtime = None
        if ucitaj:
            self.ucitaj()

    def ucitaj(self):
        self.redovi, self.dodatne = [], []
        if not os.path.isfile(self.putanja):
            self.mtime = None
            return
        redovi, zaglavlje = procitaj_csv(self.putanja)
        self.dodatne = [k for k in zaglavlje if k not in KOLONE and k not in ALIAS]
        self.redovi = redovi
        self.mtime = os.path.getmtime(self.putanja)

    def sacuvaj(self):
        zagl = KOLONE + self.dodatne
        tmp = self.putanja + ".tmp"
        with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([EN_KOL.get(k, k) for k in zagl])
            for r in self.redovi:            # redosled upisa se cuva
                vr = dict(r)
                vr["status"] = STATUS_EN.get(vr.get("status", ""), vr.get("status", ""))
                vr["pol"] = {"z": "f"}.get(vr.get("pol", ""), vr.get("pol", ""))
                w.writerow([vr.get(k, "") for k in zagl])
        os.replace(tmp, self.putanja)
        self.mtime = os.path.getmtime(self.putanja)


# column names in the saved CSV file (English) <-> internal names
EN_KOL = {"id": "id", "markica": "ear_tag", "vrsta": "species", "rasa": "breed", "pol": "sex",
          "datum_rodjenja": "birth_date", "otac": "sire", "majka": "dam", "status": "status", "stado": "herd",
          "poreklo": "origin", "poreklo_od": "origin_from", "stare_markice": "old_ear_tags", "uneto": "added"}
EN_IST = {"grlo": "animal", "datum": "date", "dogadjaj": "event", "opis": "description", "partner": "partner",
          "vrednost": "value", "automatski": "automatic", "uneto": "added", "ishod": "outcome", "ter": "heat",
          "karenca_meso": "withdrawal_meat_days", "karenca_mleko": "withdrawal_milk_days"}
EN_POD = {"grlo": "animal", "datum": "date", "tip": "type", "tekst": "text", "ponavljanje": "repeat_days",
          "gotovo": "done", "uneto": "added"}
STATUS_EN = {"aktivno": "active", "neaktivno": "inactive", "obrisano": "deleted"}
ALIAS = {"oznaka": "id", "grlo": "id", "usna_markica": "markica", "usni_broj": "markica",
         "broj_markice": "markica", "rodjen": "datum_rodjenja", "rodjeno": "datum_rodjenja",
         "datum": "datum_rodjenja", "tag": "id", "animal": "id", "name": "id", "eartag": "markica",
         "ear_tag_number": "markica", "father": "otac", "mother": "majka", "date_of_birth": "datum_rodjenja",
         "dob": "datum_rodjenja", "birthdate": "datum_rodjenja", "born": "datum_rodjenja", "gender": "pol",
         **{v: k for k, v in EN_KOL.items()}}
POL_MAP = {"m": "m", "z": "z", "ž": "z", "f": "z", "b": "m", "k": "z"}


def norm_rasa(r):
    """Rase se pišu malim slovima, osim oznake hibrida: 'hibrid L×J'."""
    r = " ".join(str(r or "").split())
    if r.lower().startswith("hybrid"):
        kod = r[6:].strip().upper().replace("X", "×").replace("*", "×").replace(" ", "")
        return f"hybrid {kod}" if kod else "hybrid"
    return r.lower()


def kod_rase(rasa, druga=""):
    """Oznaka rase u nazivu hibrida: prvo slovo (dva, ako obe rase počinju istim slovom).
    Hibrid kao roditelj ulazi u zagradi: D×(L×J)."""
    rasa = norm_rasa(rasa)
    if rasa.startswith("hybrid "):
        return f"({rasa[7:]})"
    if not rasa:
        return "?"
    d = norm_rasa(druga)
    if d and not d.startswith("hybrid") and d[0] == rasa[0] and d != rasa:
        return rasa[:2].capitalize()
    return rasa[0].upper()


def rasa_potomka(rasa_oca, rasa_majke):
    """Ista rasa -> ta rasa; različite -> 'hibrid O×M' (prvo otac, pa majka); jedna nepoznata -> poznata."""
    ro, rm = norm_rasa(rasa_oca), norm_rasa(rasa_majke)
    if not ro or ro == rm:
        return rm
    if not rm:
        return ro
    return f"hybrid {kod_rase(ro, rm)}×{kod_rase(rm, ro)}"


def vrsta_norm(v):
    v = (v or "").strip().lower()
    return {"cat": "cattle", "cow": "cattle", "bul": "cattle", "gov": "cattle", "bik": "cattle", "kra": "cattle",
            "she": "sheep", "ewe": "sheep", "ram": "sheep", "ovc": "sheep", "ova": "sheep",
            "goa": "goats", "koz": "goats", "pig": "pigs", "swi": "pigs", "sow": "pigs", "boa": "pigs",
            "svi": "pigs", "hor": "horses", "mar": "horses", "kon": "horses"}.get(v[:3], v)


def procitaj_csv(putanja):
    kod = "latin-1"
    for k in ("utf-8-sig", "cp1250"):
        try:
            with open(putanja, encoding=k) as f:
                f.read()
            kod = k
            break
        except UnicodeDecodeError:
            continue
    with open(putanja, newline="", encoding=kod) as f:
        prvi = f.readline()
        f.seek(0)
        sep = ";" if prvi.count(";") >= prvi.count(",") else ","
        citac = csv.reader(f, delimiter=sep)
        try:
            zaglavlje = [kolona(k) for k in next(citac)]
        except StopIteration:
            return [], []
        zaglavlje = [ALIAS.get(k, k) if ALIAS.get(k, k) not in zaglavlje else k for k in zaglavlje]
        if "id" not in zaglavlje and "markica" not in zaglavlje:
            raise ValueError("The table must have an 'id' (tag) or 'ear_tag' column.\n"
                             f"Columns found: {', '.join(zaglavlje)}")
        redovi = []
        for red in citac:
            if not any(x.strip() for x in red):
                continue
            r = {k: (red[i].strip() if i < len(red) else "") for i, k in enumerate(zaglavlje)}
            r["id"] = r.get("id") or r.get("markica", "")
            if not r["id"]:
                continue
            r["pol"] = POL_MAP.get(r.get("pol", "").lower()[:1], "")
            r["vrsta"] = vrsta_norm(r.get("vrsta", ""))
            r["rasa"] = norm_rasa(r.get("rasa", ""))
            for p in ("otac", "majka"):
                if r.get(p, "") == "0":
                    r[p] = ""
            st = r.get("status", "").lower()
            r["status"] = ("aktivno" if st == "" or st.startswith(("akt", "act")) else
                           "obrisano" if st.startswith(("obr", "del")) else "neaktivno")
            redovi.append(r)
        return redovi, zaglavlje


# =============================================================== istorija grla

IST_KOLONE = ["grlo", "datum", "dogadjaj", "opis", "partner", "vrednost", "automatski", "uneto", "ishod",
              "ter", "karenca_meso", "karenca_mleko"]
TER = "In heat"
# dužina polnog ciklusa (dana) – kad se ženka ponovo tera ako nije ostala gravidna
TER_CIKLUS = {"cattle": 21, "sheep": 17, "goats": 21, "pigs": 21, "horses": 21}
PARENJE = "Mating / insemination"
STADO_PODMLADAK = "Young stock"
# prosečno trajanje gravidnosti u danima
# trajanje gravidnosti (dana) i koliko dana ranije je dozvoljen upis mladih (prevremeni porođaj u granicama normale)
# nosivost: ključ -> (naziv, vrsta, rase po kojima se prepoznaje, najmanje, najviše, prosek) dana
GRAVIDNOST_PROFILI = {
    "romanov": ("Romanov sheep", "sheep", ("romanov",), 138, 143, 140),
    "sheep": ("Common sheep", "sheep", (), 145, 152, 150),
    "goats": ("Goat", "goats", (), 145, 151, 150),
    "pigs": ("Pig", "pigs", (), 110, 120, 114),
    "cattle": ("Cow", "cattle", (), 275, 290, 283),
    "horses": ("Mare", "horses", (), 320, 360, 340),
}
DOGADJAJI = ["Mating / insemination", "In heat", "Birth", "Vaccination", "Treatment", "Weighing",
             "Move", "Sale", "Death", "Note"]
AUTO = "Record change"
# kartice u istoriji: naziv -> koji događaji spadaju u grupu
GRUPE = [("Matings", ("Mating / insemination", "In heat")),
         ("Treatments", ("Treatment", "Vaccination")),
         ("Births", ("Birth", "Offspring"))]


class Istorija:
    """Dnevnik dogadjaja za svako grlo, cuva se u <tablica>_istorija.csv."""

    def __init__(self, putanja_tablice):
        self.putanja = os.path.splitext(putanja_tablice)[0] + "_history.csv"
        self.zapisi = []
        if os.path.isfile(self.putanja):
            try:
                kod = "utf-8-sig"
                try:
                    open(self.putanja, encoding=kod).read()
                except UnicodeDecodeError:
                    kod = "cp1250"
                with open(self.putanja, newline="", encoding=kod) as f:
                    prvi = f.readline()
                    f.seek(0)
                    sep = ";" if prvi.count(";") >= prvi.count(",") else ","
                    for r in csv.DictReader(f, delimiter=sep):
                        r = {({v_: k_ for k_, v_ in EN_IST.items()}.get(kolona(k), kolona(k))): (v or "").strip()
                             for k, v in r.items() if k}
                        if r.get("grlo"):
                            self.zapisi.append({k: r.get(k, "") for k in IST_KOLONE})
            except OSError:
                pass

    def sacuvaj(self):
        tmp = self.putanja + ".tmp"
        with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([EN_IST.get(k, k) for k in IST_KOLONE])
            for z in self.zapisi:
                w.writerow([z.get(k, "") for k in IST_KOLONE])
        os.replace(tmp, self.putanja)

    def dodaj(self, grlo, dogadjaj, opis="", datum=None, partner="", vrednost="", automatski=False,
              karenca_meso="", karenca_mleko=""):
        z = dict(grlo=grlo, datum=datum or date.today().strftime("%d/%m/%Y"), dogadjaj=dogadjaj, opis=opis,
                 partner=partner, vrednost=vrednost, automatski="1" if automatski else "", uneto=sada(),
                 ishod="", ter="", karenca_meso=karenca_meso, karenca_mleko=karenca_mleko)
        self.zapisi.append(z)
        return z

    def za(self, grlo=None, auto=True):
        z = [x for x in self.zapisi if (grlo is None or x["grlo"] == grlo) and (auto or not x["automatski"])]
        return sorted(z, key=lambda x: (parsiraj_datum(x["datum"]) or date.min, x["uneto"]), reverse=True)

    def preimenuj(self, staro, novo):
        for z in self.zapisi:
            if z["grlo"] == staro:
                z["grlo"] = novo
            if z["partner"] == staro:
                z["partner"] = novo


class Podsetnici:
    """Podsetnici za grla (vakcinacija, lečenje…), čuvaju se u <tablica>_podsetnici.csv."""
    KOLONE = ["grlo", "datum", "tip", "tekst", "ponavljanje", "gotovo", "uneto"]

    def __init__(self, putanja_tablice):
        self.putanja = os.path.splitext(putanja_tablice)[0] + "_reminders.csv"
        self.zapisi = []
        if os.path.isfile(self.putanja):
            try:
                kod = "utf-8-sig"
                try:
                    open(self.putanja, encoding=kod).read()
                except UnicodeDecodeError:
                    kod = "cp1250"
                with open(self.putanja, newline="", encoding=kod) as f:
                    for r in csv.DictReader(f, delimiter=";"):
                        r = {({v_: k_ for k_, v_ in EN_POD.items()}.get(kolona(k), kolona(k))): (v or "").strip()
                             for k, v in r.items() if k}
                        if r.get("datum"):
                            self.zapisi.append({k: r.get(k, "") for k in self.KOLONE})
            except OSError:
                pass

    def sacuvaj(self):
        if not self.zapisi and not os.path.isfile(self.putanja):
            return
        tmp = self.putanja + ".tmp"
        with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([EN_POD.get(k, k) for k in self.KOLONE])
            for z in self.zapisi:
                w.writerow([z.get(k, "") for k in self.KOLONE])
        os.replace(tmp, self.putanja)

    def dodaj(self, grlo, datum, tip, tekst="", ponavljanje=0):
        z = dict(grlo=grlo or "", datum=datum, tip=tip, tekst=tekst, ponavljanje=str(ponavljanje or ""),
                 gotovo="", uneto=sada())
        self.zapisi.append(z)
        return z

    def aktivni(self, grlo=None):
        return [z for z in self.zapisi if not z["gotovo"] and (grlo is None or z["grlo"] == grlo)]

    def preimenuj(self, staro, novo):
        for z in self.zapisi:
            if z["grlo"] == staro:
                z["grlo"] = novo


# =============================================================== racunanje

class Model:
    def __init__(self, redovi):
        self.A, self.P, self.by_mark, self.warn = {}, {}, {}, []
        self._gen, self._kin = {}, {}
        for rb, r in enumerate(redovi):
            jid = r["id"].strip()
            if jid in self.A:
                self.warn.append(f"Tag {jid} is entered more than once.")
                continue
            self.A[jid] = dict(id=jid, markica=r.get("markica", ""), vrsta=r.get("vrsta") or "other",
                               rasa=r.get("rasa", ""), pol=r.get("pol") or None,
                               datum=parsiraj_datum(r.get("datum_rodjenja")),
                               aktivno=r.get("status", "aktivno") == "aktivno", stado=r.get("stado", ""),
                               obrisano=r.get("status") == "obrisano",
                               osnivac=False, _otac=r.get("otac", ""), _majka=r.get("majka", ""),
                               stare=stare_lista(r), poreklo=r.get("poreklo", "").strip(),
                               poreklo_od=r.get("poreklo_od", "").strip(), rb=rb, uneto=r.get("uneto", ""))
            if r.get("markica"):
                n = norm(r["markica"])
                if n in self.by_mark and self.by_mark[n] != jid:
                    self.warn.append(f"Ear tag {r['markica']} is entered for both {self.by_mark[n]} and {jid}.")
                self.by_mark[n] = jid
        for a in list(self.A.values()):   # stari brojevi markica i dalje nalaze grlo
            for st in a.get("stare", []):
                self.by_mark.setdefault(norm(st), a["id"])
        for a in list(self.A.values()):
            self.P[a["id"]] = (self.nadji(a["_otac"]) or (a["_otac"] or None),
                               self.nadji(a["_majka"]) or (a["_majka"] or None))
        for jid, par in list(self.P.items()):
            for i, r in enumerate(par):
                if r and r not in self.A:
                    self.A[r] = dict(id=r, markica="", vrsta=self.A[jid]["vrsta"], rasa="",
                                     pol="m" if i == 0 else "z", datum=None, aktivno=False, stado="",
                                     osnivac=True, obrisano=False, stare=[], poreklo="", poreklo_od="",
                                     rb=-1, uneto="")
                    self.P[r] = (None, None)
        for o, m in list(self.P.values()):
            if o and not self.A[o]["pol"]:
                self.A[o]["pol"] = "m"
            if m and not self.A[m]["pol"]:
                self.A[m]["pol"] = "z"
        # provere
        for jid, (o, m) in list(self.P.items()):
            if jid in (o, m):
                self.warn.append(f"{jid} is entered as its own parent.")
                self.P[jid] = (None, None)
        for jid in list(self.P):
            if jid in self.preci(jid):
                self.warn.append(f"{jid} is its own ancestor – check the parents "
                                 "(this animal's pedigree is temporarily ignored).")
                self.P[jid] = (None, None)
        for jid, (o, m) in self.P.items():
            a = self.A[jid]
            for r in (o, m):
                if not r:
                    continue
                p = self.A[r]
                if p["vrsta"] != a["vrsta"] and not p["osnivac"]:
                    self.warn.append(f"{jid} ({a['vrsta']}) has parent {r} ({p['vrsta']}) – different species.")
                if p["datum"] and a["datum"] and p["datum"] >= a["datum"]:
                    self.warn.append(f"{r} was born after its offspring {jid} – check the dates.")
            if o and self.A[o]["pol"] == "z":
                self.warn.append(f"{o} is entered as sire of {jid}, but is marked as female.")
            if m and self.A[m]["pol"] == "m":
                self.warn.append(f"{m} is entered as dam of {jid}, but is marked as male.")

    def nadji(self, t):
        t = str(t or "").strip()
        if not t:
            return None
        if t in self.A:
            return t
        n = norm(t)
        if n in self.by_mark:
            return self.by_mark[n]
        for k in self.A:
            if norm(k) == n:
                return k
        return None

    def grla(self, samo_aktivna=False, redosled="Last added first"):
        """Grla za liste i pretragu (obrisana se ne prikazuju, ali ostaju u rodoslovu).
        Podrazumevano: poslednje upisano grlo je prvo, pa sve starija."""
        g = [a for a in self.A.values() if not a["osnivac"] and not a["obrisano"]
             and (a["aktivno"] or not samo_aktivna)]
        return self.poredaj(g, redosled)

    @staticmethod
    def poredaj(g, redosled="Last added first"):
        if redosled == "By tag":
            return sorted(g, key=lambda a: sort_key(a["id"]))
        if redosled == "Youngest first":
            return sorted(g, key=lambda a: (a["datum"] or date.min, a["rb"]), reverse=True)
        kljuc = lambda a: (a["uneto"] or "", a["rb"])     # bez datuma upisa -> mesto u tablici
        return sorted(g, key=kljuc, reverse=redosled != "First added first")

    def obrisana(self):
        return self.poredaj([a for a in self.A.values() if a.get("obrisano")])

    def preci(self, x):
        rez, stek = set(), [x]
        while stek:
            for r in self.P.get(stek.pop(), (None, None)):
                if r and r not in rez:
                    rez.add(r)
                    stek.append(r)
        return rez

    def gen(self, x):
        if x is None:
            return -1
        if x not in self._gen:
            o, m = self.P.get(x, (None, None))
            self._gen[x] = 1 + max(self.gen(o), self.gen(m))
        return self._gen[x]

    def kinship(self, a, b):
        if a is None or b is None:
            return 0.0
        if a == b:
            o, m = self.P.get(a, (None, None))
            return 0.5 * (1 + self.kinship(o, m))
        ga, gb = self.gen(a), self.gen(b)
        if ga > gb or (ga == gb and a > b):
            a, b = b, a
        k = (a, b)
        if k not in self._kin:
            o, m = self.P.get(b, (None, None))
            self._kin[k] = 0.5 * (self.kinship(a, o) + self.kinship(a, m))
        return self._kin[k]

    def doprinos(self, x, predak):
        """Koliki deo gena grlo x u proseku nosi od pretka (50% od roditelja, 25% od dede...)."""
        if x is None:
            return 0.0
        if x == predak:
            return 1.0
        if not hasattr(self, "_dop"):
            self._dop = {}
        k = (x, predak)
        if k not in self._dop:
            o, m = self.P.get(x, (None, None))
            self._dop[k] = 0.5 * (self.doprinos(o, predak) + self.doprinos(m, predak))
        return self._dop[k]

    def putanje(self, x):
        """Svi preci grla kroz SVE generacije: {predak: {generacija: broj puteva}}."""
        if not hasattr(self, "_put"):
            self._put = {}
        if x in self._put:
            return self._put[x]
        rez = {}
        for p in self.P.get(x, (None, None)):
            if not p:
                continue
            rez.setdefault(p, {})
            rez[p][1] = rez[p].get(1, 0) + 1
            for a, gens in self.putanje(p).items():
                d = rez.setdefault(a, {})
                for g, n in gens.items():
                    d[g + 1] = d.get(g + 1, 0) + n
        self._put[x] = rez
        return rez

    def kompletnost(self, x, do_gen):
        """Za svaku generaciju: koliko predaka je poznato od mogucih (2, 4, 8...)."""
        nivo, rez = [x], []
        for g in range(1, do_gen + 1):
            nivo = [p for y in nivo for p in self.P.get(y, (None, None)) if p]
            rez.append((g, len(nivo), 2 ** g))
            if not nivo:
                break
        return rez

    def udeli_predaka(self, x):
        return sorted(((self.doprinos(x, p), p) for p in self.preci(x)), key=lambda t: (-t[0], sort_key(t[1])))

    def inbr(self, x):
        o, m = self.P.get(x, (None, None))
        return self.kinship(o, m)

    def r(self, a, b):
        return 2 * self.kinship(a, b) / ((1 + self.inbr(a)) * (1 + self.inbr(b))) ** 0.5

    def odnos(self, a, b):
        oa, ma = self.P.get(a, (None, None))
        ob, mb = self.P.get(b, (None, None))
        if a in (ob, mb) or b in (oa, ma):
            return "parent–offspring"
        if oa and ma and oa == ob and ma == mb:
            return "full sibling"
        if (oa and oa == ob) or (ma and ma == mb):
            return "half sibling"
        if a in self.preci(b) or b in self.preci(a):
            return "ancestor and descendant"
        r = self.r(a, b)
        return ("closely related" if r >= 0.2 else "distant relatives" if r >= 0.06
                else "very distantly related" if r > 0 else "unrelated")

    def pol_naziv(self, x):
        a = self.A[x]
        n = nazivi(a["vrsta"])
        return {"m": n[0], "z": n[1]}.get(a["pol"], "sex unknown")


def procena(f, prag):
    if f == 0:
        return "excellent", OK
    if f < prag / 2:
        return "good", OK
    if f <= prag:
        return "acceptable", WARN
    if f < 0.125:
        return "risky", BAD
    return "not recommended", BAD


# =============================================================== UI pomocnici

def izaberi_font():
    fam = set(tkfont.families())
    for f in ("Segoe UI Variable Text", "Segoe UI", "SF Pro Text", "Helvetica Neue", "Ubuntu", "Cantarell",
              "DejaVu Sans"):
        if f in fam:
            return f
    return "TkDefaultFont"


def ikonica_markice(boja):
    """Mala slika usne markice (za tabele)."""
    w, h = 16, 18
    img = tk.PhotoImage(width=w, height=h)
    for y in range(h):
        xs = []
        for x in range(w):
            inside = 2 <= x <= 13 and 1 <= y <= 16
            if y < 3 and (x in (2, 13)):
                inside = False
            if y > 10:  # zaobljen donji deo
                cx, cy, rr = 7.5, 10, 6.2
                inside = inside and ((x - cx) ** 2 + (y - cy) ** 2 <= rr * rr)
            if (x - 7.5) ** 2 + (y - 5) ** 2 <= 3.2:
                inside = False
            if inside:
                xs.append(x)
        for x in xs:
            img.put(boja, (x, y))
    return img


class ScrollFrame(tk.Frame):
    def __init__(self, parent, bg=BG):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self.win = self.canvas.create_window(0, 0, window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.sb.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.win, width=e.width))
        self.bind("<Enter>", self._ukljuci)
        self.bind("<Leave>", self._iskljuci)

    def _ukljuci(self, e=None):
        self.bind_all("<MouseWheel>", self._wheel)
        self.bind_all("<Button-4>", lambda e: self._wheel(e, -1))
        self.bind_all("<Button-5>", lambda e: self._wheel(e, 1))

    def _iskljuci(self, e=None):
        for s in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.unbind_all(s)

    def _scroll(self, n):
        if self.winfo_exists() and self.winfo_ismapped():
            self.canvas.yview_scroll(n, "units")

    def _wheel(self, e, smer=None):
        # ne skroluj stranu kad je mis iznad tabele (tabela skroluje sama)
        if isinstance(e.widget, ttk.Treeview) or not self.winfo_exists():
            return
        if smer is None:
            smer = int(-e.delta / 120) if abs(e.delta) >= 120 else (-1 if e.delta > 0 else 1)
        self._scroll(smer)


# =============================================================== aplikacija

class App:
    def __init__(self, root):
        self.root = root
        self.pod = self.ucitaj_podesavanja()
        self.prag = self.pod.get("prag", 0.0625)
        self.F = izaberi_font()
        root.title("Farmer")
        self.logo = {}
        try:
            self.logo = {k: tk.PhotoImage(data=v) for k, v in LOGO.items()}
        except tk.TclError:
            pass
        # window icon: first the .ico, then the built-in logo ON TOP of it
        ico = os.path.join(APP_DIR, "farmer.ico")
        if not os.path.exists(ico): ico = os.path.join(getattr(sys, "_MEIPASS", APP_DIR), "farmer.ico")
        if sys.platform.startswith("win") and os.path.isfile(ico):
            try:
                root.iconbitmap(default=ico)
            except tk.TclError:
                pass
        if self.logo:
            try:
                root.iconphoto(True, self.logo[64], self.logo[32], self.logo[16])
            except tk.TclError:
                pass
        # Windows: set the title bar icon directly through the system (most reliable)
        if sys.platform.startswith("win") and os.path.isfile(ico):
            root.after(150, lambda: set_win_icon(root, ico))
        root.configure(bg=BG)
        root.geometry(self.pod.get("prozor", "1240x800"))
        root.minsize(980, 620)
        self.stilovi()
        self.img_tag = ikonica_markice(TAG)
        self.img_tag_siv = ikonica_markice("#B9C2BB")

        putanja = self.pod.get("tablica") or os.path.join(APP_DIR, "herd.csv")
        if not os.path.isdir(os.path.dirname(putanja) or "."):
            putanja = os.path.join(APP_DIR, "herd.csv")
        try:
            self.tab = Tablica(putanja)
        except Exception as e:
            messagebox.showerror("The table cannot be loaded", str(e))
            novi = os.path.join(APP_DIR, "herd.csv")
            if putanja != novi and os.path.isfile(novi):
                try:
                    self.tab = Tablica(novi)
                except Exception:
                    self.tab = Tablica(os.path.join(APP_DIR, "herd_new.csv"), ucitaj=False)
            else:
                self.tab = Tablica(os.path.join(APP_DIR, "herd_new.csv" if putanja == novi else "herd.csv"),
                                   ucitaj=False)
        self.model = Model(self.tab.redovi)
        self.ist = Istorija(self.tab.putanja)
        self.pods = Podsetnici(self.tab.putanja)
        self.strana = "grla"
        self.ui = {"q": "", "pol": "", "poreklo": "", "samo_akt": True, "par_tab": "partner", "partner": "", "ista_rasa": False,
                   "plan_max": 0, "sA": "", "sB": ""}
        self.karton_win = None

        self.shell()
        self.prikazi("grla")
        root.after(700, self.obavestenja_na_startu)
        root.after(1500, self.dnevna_kopija)
        root.bind("<FocusIn>", self.proveri_promenu)
        root.protocol("WM_DELETE_WINDOW", self.zatvori)

    # ------------------------------------------------------------ podesavanja
    def ucitaj_podesavanja(self):
        try:
            with open(PODESAVANJA, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return {}

    def sacuvaj_podesavanja(self):
        self.pod.update(prag=self.prag, tablica=self.tab.putanja)
        try:
            with open(PODESAVANJA, "w", encoding="utf-8") as f:
                json.dump(self.pod, f, ensure_ascii=False, indent=1)
        except OSError:
            pass

    def zatvori(self):
        try:
            self.napravi_kopiju("automatic", date.today().strftime("%Y-%m-%d"))
        except Exception:
            pass
        if self.root.state() == "normal":
            self.pod["prozor"] = self.root.geometry()
        self.sacuvaj_podesavanja()
        self.root.destroy()

    # ------------------------------------------------------------ stil
    def f(self, size=11, weight="normal"):
        return (self.F, size, weight)

    def stilovi(self):
        s = ttk.Style()
        s.theme_use("clam")
        F = self.F
        s.configure(".", background=BG, foreground=INK, font=(F, 10), bordercolor=LINE,
                    lightcolor=LINE, darkcolor=LINE, troughcolor=SURF2, focuscolor=BRAND)
        for name, bg, fg, hov, bd in (("Primary", BRAND, "#FFFFFF", BRAND_H, BRAND),
                                      ("Secondary", SURF, INK, SURF2, LINE),
                                      ("Danger", SURF, BAD, BAD_BG, SURF)):
            s.configure(f"{name}.TButton", background=bg, foreground=fg, font=(F, 10, "bold"),
                        padding=(16, 8), borderwidth=1, bordercolor=bd, lightcolor=bg, darkcolor=bg,
                        focusthickness=0, relief="solid")
            s.map(f"{name}.TButton", background=[("active", hov), ("disabled", SURF2)],
                  lightcolor=[("active", hov)], darkcolor=[("active", hov)],
                  foreground=[("disabled", INK2)])
        s.configure("TEntry", fieldbackground=SURF, padding=(10, 7), bordercolor=LINE,
                    lightcolor=SURF, darkcolor=SURF, insertcolor=INK)
        s.map("TEntry", bordercolor=[("focus", BRAND)], lightcolor=[("focus", BRAND)])
        s.configure("TCombobox", fieldbackground=SURF, background=SURF, padding=(10, 6), arrowsize=14,
                    bordercolor=LINE, lightcolor=SURF, darkcolor=SURF, arrowcolor=INK2)
        s.map("TCombobox", fieldbackground=[("readonly", SURF)], bordercolor=[("focus", BRAND)],
              selectbackground=[("readonly", SURF)], selectforeground=[("readonly", INK)])
        self.root.option_add("*TCombobox*Listbox.font", (F, 10))
        self.root.option_add("*TCombobox*Listbox.selectBackground", BRAND)
        s.configure("TCheckbutton", background=BG, font=(F, 10), indicatorcolor=SURF,
                    indicatorbackground=SURF)
        s.map("TCheckbutton", indicatorcolor=[("selected", BRAND)], background=[("active", BG)])
        s.configure("Card.TCheckbutton", background=SURF)
        s.map("Card.TCheckbutton", background=[("active", SURF)])
        s.configure("Card.TRadiobutton", background=SURF, font=(F, 10), indicatorcolor=SURF)
        s.map("Card.TRadiobutton", indicatorcolor=[("selected", BRAND)], background=[("active", SURF)])
        s.configure("Treeview", background=SURF, fieldbackground=SURF, foreground=INK, rowheight=36,
                    borderwidth=0, font=(F, 10))
        s.map("Treeview", background=[("selected", "#D5E6DA")], foreground=[("selected", INK)])
        s.configure("Treeview.Heading", background=SURF2, foreground=INK2, font=(F, 9, "bold"),
                    relief="flat", padding=(8, 8), borderwidth=0)
        s.map("Treeview.Heading", background=[("active", SURF2)])
        s.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        s.configure("Green.Horizontal.TProgressbar", troughcolor=SURF2, background=BRAND, thickness=8,
                    borderwidth=0, lightcolor=BRAND, darkcolor=BRAND)
        s.configure("Vertical.TScrollbar", background=SURF2, troughcolor=BG, bordercolor=BG, arrowcolor=INK2,
                    lightcolor=SURF2, darkcolor=SURF2, gripcount=0)

    # ------------------------------------------------------------ okvir
    def shell(self):
        side = tk.Frame(self.root, bg=SURF, width=230, highlightthickness=0)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        tk.Frame(self.root, bg=LINE, width=1).pack(side="left", fill="y")

        brand = tk.Frame(side, bg=SURF)
        brand.pack(fill="x", padx=14, pady=(22, 26))
        if self.logo.get(36):
            tk.Label(brand, image=self.logo[36], bg=SURF, bd=0).pack(side="left")
        else:
            c = tk.Canvas(brand, width=34, height=38, bg=SURF, highlightthickness=0)
            c.pack(side="left")
            self.crtaj_markicu(c, 2, 2, 30, 34)
        t = tk.Frame(brand, bg=SURF)
        t.pack(side="left", padx=10)
        tk.Label(t, text="Farmer", bg=SURF, fg=INK, font=self.f(12, "bold")).pack(anchor="w")
        tk.Label(t, text="kinship & mating", bg=SURF, fg=INK2, font=self.f(9)).pack(anchor="w")

        self.nav = {}
        for key, txt, ic in (("grla", "Animals", "☰"), ("parenje", "Mating", "⚭"),
                             ("srodstvo", "Relationship", "⑂"), ("istorija", "History", "◷"),
                             ("obavestenja", "Notifications", "◉"), ("pregled", "Overview", "▥")):
            row = tk.Frame(side, bg=SURF, cursor="hand2")
            row.pack(fill="x", padx=12, pady=1)
            li = tk.Label(row, text=ic, bg=SURF, fg=INK2, font=self.f(12), width=2)
            li.pack(side="left", padx=(10, 4), pady=9)
            lt = tk.Label(row, text=txt, bg=SURF, fg=INK2, font=self.f(11, "bold"), anchor="w")
            lt.pack(side="left", fill="x", expand=True)
            for w in (row, li, lt):
                w.bind("<Button-1>", lambda e, k=key: self.prikazi(k))
                w.bind("<Enter>", lambda e, k=key: self.nav_hover(k, True))
                w.bind("<Leave>", lambda e, k=key: self.nav_hover(k, False))
            self.nav[key] = (row, li, lt)

        foot = tk.Frame(side, bg=SURF)
        foot.pack(side="bottom", fill="x", padx=18, pady=18)
        self.lbl_fajl = tk.Label(foot, text="", bg=SURF, fg=INK2, font=self.f(9), anchor="w", justify="left",
                                 wraplength=190)
        self.lbl_fajl.pack(anchor="w")
        lnk = tk.Label(foot, text="Change table…", bg=SURF, fg=BRAND, font=self.f(9, "bold"), cursor="hand2")
        lnk.pack(anchor="w", pady=(4, 0))
        lnk.bind("<Button-1>", lambda e: self.promeni_tablicu())
        self.lbl_kopija = tk.Label(foot, text="", bg=SURF, fg=INK2, font=self.f(9), anchor="w", cursor="hand2")
        self.lbl_kopija.pack(anchor="w", pady=(8, 0))
        self.lbl_kopija.bind("<Button-1>", lambda e: self.prikazi("pregled"))
        self.osvezi_fajl_label()
        self.osvezi_kopija_label()

        self.main = tk.Frame(self.root, bg=BG)
        self.main.pack(side="left", fill="both", expand=True)
        self.toast_lbl = None

    def osvezi_fajl_label(self):
        self.lbl_fajl.config(text=f"Data: {os.path.basename(self.tab.putanja)}")

    def nav_hover(self, k, on):
        row, li, lt = self.nav[k]
        bg = SURF2 if (on or k == self.strana) else SURF
        for w in (row, li, lt):
            w.config(bg=bg)

    def crtaj_markicu(self, c, x, y, w, h, boja=TAG, rupa=SURF):
        r = w * 0.18
        c.create_polygon(self.zaobljeno(x, y, x + w, y + h * 0.62, r) + [x + w, y + h * 0.62],
                         fill=boja, outline=boja, smooth=True)
        c.create_arc(x, y + h - w, x + w, y + h, start=180, extent=180, fill=boja, outline=boja)
        c.create_rectangle(x, y + h * 0.4, x + w, y + h - w / 2, fill=boja, outline=boja)
        cx, cy, rr = x + w / 2, y + h * 0.22, w * 0.1
        c.create_oval(cx - rr, cy - rr, cx + rr, cy + rr, fill=rupa, outline="#B8A04A")
        c.create_line(x + w * 0.22, y + h * 0.52, x + w * 0.78, y + h * 0.52, fill=TAG_INK, width=2.4,
                      capstyle="round")
        c.create_line(x + w * 0.22, y + h * 0.7, x + w * 0.55, y + h * 0.7, fill=TAG_INK, width=2.4,
                      capstyle="round")

    @staticmethod
    def zaobljeno(x1, y1, x2, y2, r):
        return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
                x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]

    def cip(self, parent, text, cmd=None, siv=False, bg=SURF, big=False):
        """Ciparska 'usna markica' sa oznakom grla."""
        fnt = self.f(16 if big else 10, "bold")
        fo = tkfont.Font(font=fnt)
        tw = fo.measure(text)
        h = 40 if big else 26
        pl, pr = (34, 16) if big else (22, 10)
        w = tw + pl + pr
        c = tk.Canvas(parent, width=w, height=h, bg=bg, highlightthickness=0, cursor="hand2" if cmd else "")
        boja = "#DDE3DC" if siv else TAG
        r = 9 if big else 7
        c.create_polygon(self.zaobljeno(1, 1, w - 1, h - 1, r), fill=boja, outline=boja, smooth=True)
        hr = 5 if big else 3.5
        hx = 17 if big else 11
        c.create_oval(hx - hr, h / 2 - hr, hx + hr, h / 2 + hr, fill=bg, outline="#A8963F" if not siv else "#AAB3AC")
        c.create_text(pl, h / 2, text=text, anchor="w", font=fnt, fill=TAG_INK if not siv else INK2)
        if cmd:
            c.bind("<Button-1>", lambda e: cmd())
        return c

    def btn(self, parent, text, cmd, kind="Secondary"):
        return ttk.Button(parent, text=text, command=cmd, style=f"{kind}.TButton", cursor="hand2")

    def card(self, parent, padx=20, pady=18):
        outer = tk.Frame(parent, bg=LINE)
        inner = tk.Frame(outer, bg=SURF)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        body = tk.Frame(inner, bg=SURF)
        body.pack(fill="both", expand=True, padx=padx, pady=pady)
        return outer, body

    def lbl(self, parent, text, size=10, weight="normal", fg=INK, bg=None, **kw):
        return tk.Label(parent, text=text, font=self.f(size, weight), fg=fg,
                        bg=bg or parent.cget("bg"), **kw)

    def badge(self, parent, text, fg, bg=None):
        bgc = {OK: OK_BG, WARN: WARN_BG, BAD: BAD_BG}.get(fg, SURF2)
        return tk.Label(parent, text=f" {text} ", font=self.f(9, "bold"), fg=fg, bg=bgc, padx=6, pady=1)

    def toast(self, text):
        if self.toast_lbl and self.toast_lbl.winfo_exists():
            self.toast_lbl.destroy()
        t = tk.Label(self.main, text=text, bg=INK, fg="#FFFFFF", font=self.f(10, "bold"), padx=18, pady=10)
        t.place(relx=0.5, rely=1.0, y=-24, anchor="s")
        self.toast_lbl = t
        self.root.after(2600, lambda: t.winfo_exists() and t.destroy())

    def stat(self, parent, broj, tekst, boja=INK):
        o, b = self.card(parent, 16, 12)
        self.lbl(b, broj, 22, "bold", fg=boja).pack(anchor="w")
        self.lbl(b, tekst, 9, fg=INK2).pack(anchor="w")
        return o

    def tabela(self, parent, kolone, visina=10, tree=True):
        wrap = tk.Frame(parent, bg=LINE)
        inner = tk.Frame(wrap, bg=SURF)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        cols = [k for k, *_ in kolone[1:]] if tree else [k for k, *_ in kolone]
        tv = ttk.Treeview(inner, columns=cols, show="tree headings" if tree else "headings", height=visina,
                          selectmode="extended")
        sb = ttk.Scrollbar(inner, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        for i, (k, naslov, w, *anc) in enumerate(kolone):
            cid = "#0" if (tree and i == 0) else k
            tv.heading(cid, text=naslov, anchor=anc[0] if anc else "w")
            tv.column(cid, width=w, minwidth=40, anchor=anc[0] if anc else "w", stretch=True)
        tv.tag_configure("odd", background="#FAFBF9")
        tv.tag_configure("off", foreground="#98A39B")
        tv.tag_configure("bad", foreground=BAD)
        tv.tag_configure("warn", foreground=WARN)
        tv.tag_configure("ok", foreground=OK)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return wrap, tv

    # ------------------------------------------------------------ podaci
    def promena(self, poruka=None):
        try:
            self.ist.sacuvaj()
            self.pods.sacuvaj()
        except OSError:
            pass
        try:
            self.tab.sacuvaj()
        except PermissionError:
            messagebox.showerror("Not saved",
                                 "The table is open in Excel or another program.\n"
                                 "Close it and try again.")
            self.tab.ucitaj()
        self.model = Model(self.tab.redovi)
        self.prikazi(self.strana)
        if self.karton_win and self.karton_win.winfo_exists():
            self.karton(self.karton_win.grlo)
        if poruka:
            self.toast(poruka)

    def proveri_promenu(self, e=None):
        if e is not None and e.widget is not self.root:
            return
        p = self.tab.putanja
        if os.path.isfile(p) and self.tab.mtime and os.path.getmtime(p) != self.tab.mtime:
            try:
                self.tab.ucitaj()
            except Exception as ex:
                messagebox.showerror("Error in table", str(ex))
                return
            self.model = Model(self.tab.redovi)
            self.prikazi(self.strana)
            self.toast("The table was reloaded (changed in Excel).")

    def promeni_tablicu(self):
        p = filedialog.askopenfilename(title="Choose the animals table", initialdir=APP_DIR,
                                       filetypes=[("CSV table", "*.csv"), ("All files", "*.*")])
        if not p:
            return
        try:
            t = Tablica(p)
        except Exception as e:
            messagebox.showerror("The table cannot be loaded", str(e))
            return
        self.tab = t
        self.ist = Istorija(t.putanja)
        self.pods = Podsetnici(t.putanja)
        self.model = Model(t.redovi)
        self.sacuvaj_podesavanja()
        self.osvezi_fajl_label()
        self.prikazi(self.strana)
        self.toast(f"Loaded {len(t.redovi)} animals.")

    def uvezi(self):
        p = filedialog.askopenfilename(title="Import table from Excel (CSV)",
                                       filetypes=[("CSV table", "*.csv"), ("All files", "*.*")])
        if not p:
            return
        try:
            redovi, _ = procitaj_csv(p)
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
            return
        po_id = {r["id"]: r for r in self.tab.redovi}
        nov = izm = 0
        for r in redovi:
            r = {k: r.get(k, "") for k in KOLONE}
            if r["id"] in po_id:
                if not r["uneto"]:
                    r.pop("uneto")
                po_id[r["id"]].update(r)
                izm += 1
            else:
                r["uneto"] = r.get("uneto") or sada()
                self.tab.redovi.append(r)
                self.ist.dodaj(r["id"], AUTO, f"Added by import from table {os.path.basename(p)}", automatski=True)
                po_id[r["id"]] = r
                nov += 1
        self.promena(f"Imported: {nov} new" + (f", {izm} changed" if izm else "") + " animals.")

    def otvori_excel(self):
        if not os.path.isfile(self.tab.putanja):
            self.tab.sacuvaj()
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.tab.putanja)
            elif sys.platform == "darwin":
                os.system(f'open "{self.tab.putanja}"')
            else:
                os.system(f'xdg-open "{self.tab.putanja}" &')
            self.toast("When you save changes in Excel, the program will reload them.")
        except OSError:
            messagebox.showinfo("Table", self.tab.putanja)

    # ------------------------------------------------------------ strane
    def prikazi(self, strana):
        self.strana = strana
        for k in self.nav:
            row, li, lt = self.nav[k]
            akt = k == strana or (k == "grla" and strana == "obrisana")
            for w in (row, li, lt):
                w.config(bg=SURF2 if akt else SURF)
            lt.config(fg=INK if akt else INK2)
            li.config(fg=BRAND if akt else INK2)
        self.osvezi_nav_brojac()
        for w in self.main.winfo_children():
            if w is not self.toast_lbl:
                w.destroy()
        getattr(self, "s_" + strana)()

    def zaglavlje(self, parent, naslov, podnaslov="", akcije=()):
        h = tk.Frame(parent, bg=BG)
        h.pack(fill="x", pady=(0, 18))
        desno = tk.Frame(h, bg=BG)
        desno.pack(side="right", anchor="s")
        for txt, cmd, kind in reversed(akcije):
            self.btn(desno, txt, cmd, kind).pack(side="right", padx=(8, 0))
        levo = tk.Frame(h, bg=BG)
        levo.pack(side="left", fill="x", expand=True)
        self.lbl(levo, naslov, 24, "bold").pack(anchor="w")
        if podnaslov:
            self.lbl(levo, podnaslov, 10, fg=INK2, justify="left",
                     wraplength=620 if not akcije else 300).pack(anchor="w", pady=(2, 0))
        return h

    def prazno(self, parent):
        o, b = self.card(parent, 30, 50)
        o.pack(fill="x")
        if self.logo.get(64):
            tk.Label(b, image=self.logo[64], bg=SURF, bd=0).pack()
        else:
            c = tk.Canvas(b, width=56, height=64, bg=SURF, highlightthickness=0)
            c.pack()
            self.crtaj_markicu(c, 4, 2, 48, 58)
        self.lbl(b, "The herd is still empty", 16, "bold").pack(pady=(14, 4))
        self.lbl(b, "Add animals one by one or import an existing table from Excel.\n"
                    "For each animal enter the sire and dam if you know them.", 10, fg=INK2, justify="center").pack()
        r = tk.Frame(b, bg=SURF)
        r.pack(pady=(18, 0))
        self.btn(r, "+  Add first animal", lambda: self.forma(), "Primary").pack(side="left", padx=4)
        self.btn(r, "Import table from Excel", self.uvezi).pack(side="left", padx=4)

    def strana_okvir(self, scroll=True):
        if scroll:
            sf = ScrollFrame(self.main)
            sf.pack(fill="both", expand=True)
            body = tk.Frame(sf.inner, bg=BG)
            body.pack(fill="both", expand=True, padx=36, pady=28)
            return body
        body = tk.Frame(self.main, bg=BG)
        body.pack(fill="both", expand=True, padx=36, pady=28)
        return body

    # ---------- Grla
    def s_grla(self):
        body = self.strana_okvir(scroll=False)
        sva = self.model.grla()
        akt = sum(a["aktivno"] for a in sva)
        self.zaglavlje(body, "Animals", f"{akt} in herd, {len(sva)} recorded in total."
                       if sva else "Records of all animals with their parents.",
                       [("Partner search", self.idi_na_pretragu, "Secondary"),
                        ("Import", self.uvezi, "Secondary"), ("Open in Excel", self.otvori_excel, "Secondary"),
                        ("+  New animal", lambda: self.forma(), "Primary")])
        if not sva:
            self.prazno(body)
            return
        tb = tk.Frame(body, bg=BG)
        tb.pack(fill="x", pady=(0, 12))
        self.q_var = tk.StringVar(value=self.ui["q"])
        e = ttk.Entry(tb, textvariable=self.q_var, font=self.f(10), width=26)
        e.pack(side="left")
        ph = self.lbl(tb, "⌕  Search by tag, ear tag, breed…", 10, fg="#98A39B", bg=SURF)
        def ph_show(*_):
            if self.q_var.get():
                ph.place_forget()
            else:
                ph.place(in_=e, x=10, rely=0.5, anchor="w")
        ph.bind("<Button-1>", lambda ev: e.focus_set())
        ph_show()
        self.pol_var = tk.StringVar(value={"": "All sexes", "z": "Females", "m": "Males"}[self.ui["pol"]])
        cb = ttk.Combobox(tb, textvariable=self.pol_var, values=["All sexes", "Females", "Males"], width=10,
                          state="readonly", font=self.f(10))
        cb.pack(side="left", padx=10)
        self.por_var = tk.StringVar(value=self.ui.get("poreklo") or "Any origin")
        cbp = ttk.Combobox(tb, textvariable=self.por_var, values=["Any origin"] + POREKLA, width=15,
                           state="readonly", font=self.f(10))
        cbp.pack(side="left", padx=(0, 10))
        cbp.bind("<<ComboboxSelected>>", lambda ev: self.grla_filter())
        self.akt_var = tk.BooleanVar(value=self.ui["samo_akt"])
        ttk.Checkbutton(tb, text="Only in herd", variable=self.akt_var,
                        command=lambda: self.grla_filter()).pack(side="left", padx=6)
        self.red_var = tk.StringVar(value=self.ui.get("redosled") or REDOSLEDI[0])
        cbr = ttk.Combobox(tb, textvariable=self.red_var, values=REDOSLEDI, state="readonly", width=21,
                           font=self.f(10))
        cbr.pack(side="right")
        cbr.bind("<<ComboboxSelected>>", lambda ev: self.grla_filter())
        self.lbl(tb, "Order:", 9, "bold", fg=INK2).pack(side="right", padx=(0, 6))

        wrap, tv = self.tabela(body, [("id", "Animal", 105), ("markica", "Ear tag", 125), ("pol", "Sex", 70),
                                      ("rasa", "Breed", 130), ("starost", "Age", 90), ("otac", "Sire", 100),
                                      ("majka", "Dam", 100), ("upis", "Added", 95), ("f", "Inbreeding", 90, "e")],
                               visina=14)
        wrap.pack(fill="both", expand=True)
        self.tv_grla = tv
        tv.bind("<Double-1>", lambda ev: self.otvori_iz(tv))
        tv.bind("<Return>", lambda ev: self.otvori_iz(tv))
        tv.bind("<Delete>", lambda ev: self.obrisi_grla([i for i in tv.selection() if i in self.model.A]))
        tv.bind("<Button-3>", lambda ev: self.meni_grla(tv, ev))
        tv.bind("<Button-2>", lambda ev: self.meni_grla(tv, ev))
        dno = tk.Frame(body, bg=BG)
        dno.pack(fill="x", pady=(8, 0))
        self.lbl(dno, "Double-click opens the card, right-click gives more options. To delete, select an animal (or several with Ctrl) and press Delete.",
                 9, fg=INK2).pack(side="left")
        n_obr = len(self.model.obrisana())
        if n_obr:
            lnk = self.lbl(dno, f"Deleted animals ({n_obr})", 9, "bold", fg=BRAND, cursor="hand2")
            lnk.pack(side="right")
            lnk.bind("<Button-1>", lambda e: self.prikazi("obrisana"))
        self.q_var.trace_add("write", lambda *_: (ph_show(), self.grla_filter()))
        cb.bind("<<ComboboxSelected>>", lambda ev: self.grla_filter())
        self.grla_filter()

    def grla_filter(self):
        self.ui["q"] = self.q_var.get()
        self.ui["pol"] = {"Females": "z", "Males": "m"}.get(self.pol_var.get(), "")
        self.ui["poreklo"] = "" if self.por_var.get() == "Any origin" else self.por_var.get()
        self.ui["samo_akt"] = self.akt_var.get()
        self.ui["redosled"] = self.red_var.get()
        tv, m, q = self.tv_grla, self.model, norm(self.ui["q"])
        tv.delete(*tv.get_children())
        i = 0
        for a in m.grla(redosled=self.ui["redosled"]):
            if self.ui["samo_akt"] and not a["aktivno"]:
                continue
            if self.ui["poreklo"] and norm(a["poreklo"]) != norm(self.ui["poreklo"]):
                continue
            if self.ui["pol"] and a["pol"] != self.ui["pol"]:
                continue
            if q and not (any(q in norm(a[k]) for k in ("id", "markica", "rasa", "stado", "poreklo_od"))
                          or any(q in norm(x) for x in a["stare"])):
                continue
            o, mj = m.P[a["id"]]
            f = m.inbr(a["id"])
            tags = ["odd"] if i % 2 else []
            if not a["aktivno"]:
                tags.append("off")
            elif f > self.prag:
                tags.append("bad")
            if self.karenca(a["id"]):
                tags.append("kar")
            tv.tag_configure("kar", foreground=WARN)
            tv.insert("", "end", iid=a["id"], text="  " + a["id"] + ("  ✚" if "kar" in tags else ""), image=self.img_tag,
                      values=(a["markica"] or "–", m.pol_naziv(a["id"]), a["rasa"] or "–",
                              fmt_starost(starost_meseci(a["datum"])) or "–", o or "–", mj or "–",
                              fmt_upis(a["uneto"]), pct(f) + ("  ⚠" if f > self.prag else "")), tags=tags)
            i += 1
        if not i:
            tv.insert("", "end", text="", values=("", "No animal matches the search.", "", "", "", "", "", ""))

    def idi_na_pretragu(self, x=None):
        self.ui.update(par_tab="partner", partner_vise=[])
        if x:
            self.ui.update(partner=x, partner_upit=x)
        self.prikazi("parenje")

    def otvori_iz(self, tv):
        sel = tv.selection()
        if sel and sel[0] in self.model.A:
            self.karton(sel[0])

    # ---------- brisanje
    def red_za(self, jid):
        return next((r for r in self.tab.redovi if r["id"] == jid), None)

    # ---------- usne markice
    def markica_zauzeta(self, mk, red):
        """Vrati poruku ako je markica vec kod drugog grla, inace None."""
        n = norm(mk)
        for r in self.tab.redovi:
            if r is red:
                continue
            if r.get("markica") and norm(r["markica"]) == n:
                return f"Ear tag {mk} is already entered for animal {r['id']}."
            if any(norm(x) == n for x in stare_lista(r)):
                return f"Ear tag {mk} previously belonged to animal {r['id']}."
            if norm(r["id"]) == n:
                return f"{mk} is already the tag of animal {r['id']}."
        return None

    def primeni_markicu(self, red, nova, zapamti=True):
        """Promeni markicu i popravi potomke koji su roditelja upisali preko stare markice."""
        stara = (red.get("markica") or "").strip()
        nova = nova.strip()
        if stara and norm(stara) != norm(nova):
            n = norm(stara)
            for r in self.tab.redovi:
                if r is red:
                    continue
                for p in ("otac", "majka"):
                    if r.get(p) and norm(r[p]) == n and r[p] != red["id"]:
                        r[p] = red["id"]
            if zapamti:
                st = stare_lista(red)
                if all(norm(x) != n for x in st):
                    st.append(stara)
                red["stare_markice"] = ", ".join(st)
        red["stare_markice"] = ", ".join(x for x in stare_lista(red) if norm(x) != norm(nova))
        if norm(stara) != norm(nova):
            self.ist.dodaj(red["id"], AUTO, f"Ear tag changed: {stara or '–'} → {nova or '–'}", automatski=True)
        red["markica"] = nova

    def promena_markice(self, jid, parent=None):
        red = self.red_za(jid)
        if not red:
            return
        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title("Change ear tag")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, "Change ear tag", 16, "bold").pack(anchor="w")
        top = tk.Frame(b, bg=SURF)
        top.pack(anchor="w", pady=(14, 4))
        self.cip(top, jid).pack(side="left")
        self.lbl(top, f"current ear tag:  {red.get('markica') or '–'}", 10, fg=INK2).pack(side="left", padx=12)
        stare = stare_lista(red)
        if stare:
            self.lbl(b, "Previous ear tags: " + ", ".join(stare), 9, fg=INK2).pack(anchor="w")
        self.lbl(b, "New ear tag", 9, "bold", fg=INK2).pack(anchor="w", pady=(14, 4))
        var = tk.StringVar(value=red.get("markica", ""))
        e = ttk.Entry(b, textvariable=var, width=36, font=self.f(11))
        e.pack(anchor="w", fill="x")
        self.lbl(b, "e.g. UK 123456 700123. Leave empty to remove the ear tag.",
                 8, fg="#8A968D").pack(anchor="w")
        zap = tk.BooleanVar(value=True)
        ttk.Checkbutton(b, text="Remember the old ear tag (search still finds the animal by its old number)",
                        variable=zap, style="Card.TCheckbutton").pack(anchor="w", pady=(12, 0))
        err = self.lbl(b, "", 10, "bold", fg=BAD)
        err.pack(anchor="w", pady=(8, 0))

        def sacuvaj(ev=None):
            nova = var.get().strip()
            if norm(nova) == norm(red.get("markica")) and nova == (red.get("markica") or ""):
                win.destroy()
                return
            if nova:
                poruka = self.markica_zauzeta(nova, red)
                if poruka:
                    err.config(text=poruka)
                    return
            stara = red.get("markica") or "–"
            self.primeni_markicu(red, nova, zapamti=zap.get())
            win.grab_release()
            win.destroy()
            self.promena(f"{jid}: ear tag {stara} → {nova or '–'}")

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(18, 0))
        self.btn(fb, "Save", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        win.bind("<Return>", sacuvaj)
        win.bind("<Escape>", lambda ev: win.destroy())
        win.update_idletasks()
        r = parent or self.root
        win.geometry(f"+{r.winfo_rootx() + max(0, (r.winfo_width() - win.winfo_width()) // 2)}"
                     f"+{r.winfo_rooty() + max(20, (r.winfo_height() - win.winfo_height()) // 3)}")
        e.focus_set()
        e.select_range(0, "end")

    # ---------- stado
    def sva_stada(self):
        return sorted({r.get("stado", "").strip() for r in self.tab.redovi if r.get("stado", "").strip()},
                      key=sort_key)

    def promena_stada(self, ids, parent=None):
        ids = [i for i in ids if self.red_za(i)]
        if not ids:
            return
        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title("Move to another herd")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, "Move to another herd", 16, "bold").pack(anchor="w")
        top = tk.Frame(b, bg=SURF)
        top.pack(anchor="w", pady=(12, 4))
        for i in ids[:8]:
            self.cip(top, i).pack(side="left", padx=(0, 6))
        if len(ids) > 8:
            self.lbl(top, f"+ {len(ids) - 8} more", 9, fg=INK2).pack(side="left")
        sada_st = sorted({self.red_za(i).get("stado") or "–" for i in ids})
        self.lbl(b, "Current herd: " + ", ".join(sada_st), 10, fg=INK2).pack(anchor="w")
        self.lbl(b, "New herd", 9, "bold", fg=INK2).pack(anchor="w", pady=(14, 4))
        sv = tk.StringVar(value="")
        e = ttk.Combobox(b, textvariable=sv, values=self.sva_stada(), width=34, font=self.f(11))
        e.pack(anchor="w", fill="x")
        self.lbl(b, "Choose an existing herd or type a new name.", 8, fg="#8A968D").pack(anchor="w")
        self.lbl(b, "Date of move", 9, "bold", fg=INK2).pack(anchor="w", pady=(12, 4))
        dv = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        ttk.Entry(b, textvariable=dv, width=16, font=self.f(10)).pack(anchor="w")
        zv = tk.BooleanVar(value=True)
        ttk.Checkbutton(b, text="Record the move in the animal's history", variable=zv,
                        style="Card.TCheckbutton").pack(anchor="w", pady=(12, 0))
        err = self.lbl(b, "", 10, "bold", fg=BAD)
        err.pack(anchor="w", pady=(8, 0))

        def sacuvaj(ev=None):
            novo = sv.get().strip()
            if not novo:
                err.config(text="Type or choose the new herd.")
                return
            d = parsiraj_datum(dv.get())
            if not d:
                err.config(text="The date is not valid. Enter it as 15/03/2026.")
                return
            # isti naziv kao postojece stado (bez obzira na velika slova) -> koristi postojeci
            novo = next((x for x in self.sva_stada() if x.lower() == novo.lower()), novo)
            pomereno = 0
            for i in ids:
                r = self.red_za(i)
                staro = (r.get("stado") or "").strip()
                if staro == novo:
                    continue
                r["stado"] = novo
                pomereno += 1
                if zv.get():
                    self.ist.dodaj(i, "Move", f"From herd {staro or '–'} to herd {novo}",
                                   datum=d.strftime("%d/%m/%Y"), vrednost=novo)
            win.grab_release()
            win.destroy()
            self.promena(f"Moved to herd {novo}: {pomereno} animals" if pomereno else "The animals are already in that herd.")

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(18, 0))
        self.btn(fb, "Move", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        win.bind("<Return>", sacuvaj)
        win.bind("<Escape>", lambda ev: win.destroy())
        win.update_idletasks()
        r = parent or self.root
        win.geometry(f"+{r.winfo_rootx() + max(0, (r.winfo_width() - win.winfo_width()) // 2)}"
                     f"+{r.winfo_rooty() + max(20, (r.winfo_height() - win.winfo_height()) // 3)}")
        e.focus_set()

    def preimenuj_stado(self, staro):
        win = tk.Toplevel(self.root, bg=SURF)
        win.title("Rename herd")
        win.transient(self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        n = sum(1 for r in self.tab.redovi if (r.get("stado") or "").strip() == staro)
        self.lbl(b, "Rename herd", 16, "bold").pack(anchor="w")
        self.lbl(b, f"Herd \"{staro}\" has {n} animals. The new name applies to all of them.", 10, fg=INK2).pack(
            anchor="w", pady=(8, 0))
        self.lbl(b, "New name", 9, "bold", fg=INK2).pack(anchor="w", pady=(14, 4))
        sv = tk.StringVar(value=staro)
        e = ttk.Entry(b, textvariable=sv, width=34, font=self.f(11))
        e.pack(anchor="w", fill="x")
        err = self.lbl(b, "", 10, "bold", fg=BAD)
        err.pack(anchor="w", pady=(8, 0))

        def sacuvaj(ev=None):
            novo = sv.get().strip()
            if not novo:
                err.config(text="Type a new name.")
                return
            if novo == staro:
                win.destroy()
                return
            for r in self.tab.redovi:
                if (r.get("stado") or "").strip() == staro:
                    r["stado"] = novo
                    self.ist.dodaj(r["id"], AUTO, f"Herd renamed: {staro} → {novo}", automatski=True)
            win.grab_release()
            win.destroy()
            self.promena(f"Herd \"{staro}\" is now \"{novo}\"")

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(18, 0))
        self.btn(fb, "Save", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        win.bind("<Return>", sacuvaj)
        win.bind("<Escape>", lambda ev: win.destroy())
        win.update_idletasks()
        win.geometry(f"+{self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2}"
                     f"+{self.root.winfo_rooty() + 120}")
        e.focus_set()
        e.select_range(0, "end")

    # ---------- brisanje svih grla (5 koraka)
    def brisanje_svih(self):
        n_grla = len(self.tab.redovi)
        if not n_grla:
            messagebox.showinfo("Delete all animals", "There are no animals in the records.")
            return
        n_ist = len(self.ist.zapisi)
        win = tk.Toplevel(self.root, bg=SURF)
        win.title("Delete all animals")
        win.transient(self.root)
        win.resizable(False, False)
        win.grab_set()
        st = {"korak": 1, "istorija": tk.StringVar(value="sve"), "kopija": None, "tekst": tk.StringVar(),
              "odbrojavanje": 0, "gotovo": False}
        NASLOVI = ["Warning", "What is deleted", "Backup", "Confirmation", "Deletion"]

        glava = tk.Frame(win, bg=SURF)
        glava.pack(fill="x", padx=26, pady=(22, 0))
        sadrzaj = tk.Frame(win, bg=SURF, width=560, height=290)
        sadrzaj.pack(fill="both", expand=True, padx=26, pady=(14, 0))
        sadrzaj.pack_propagate(False)
        dno = tk.Frame(win, bg=SURF)
        dno.pack(fill="x", padx=26, pady=(10, 20))

        def nacrtaj():
            for w in glava.winfo_children() + sadrzaj.winfo_children() + dno.winfo_children():
                w.destroy()
            k = st["korak"]
            self.lbl(glava, f"Step {k} of 5 – {NASLOVI[k - 1]}", 16, "bold",
                     fg=BAD if k >= 4 else INK).pack(anchor="w")
            tr = tk.Frame(glava, bg=SURF)
            tr.pack(anchor="w", pady=(10, 0))
            for i in range(1, 6):
                boja = BAD if i <= k and k >= 4 else BRAND if i <= k else SURF2
                tk.Frame(tr, bg=boja, width=100, height=6).pack(side="left", padx=(0, 6))
            [k1, k2, k3, k4, k5][k - 1]()

        def dugmad(dalje_txt="Next  ›", dalje=None, dalje_stil="Primary", nazad=True, stanje="normal"):
            self.btn(dno, "Cancel", win.destroy).pack(side="left")
            b = self.btn(dno, dalje_txt, dalje or sledeci, dalje_stil)
            b.pack(side="right")
            if stanje == "disabled":
                b.state(["disabled"])
            if nazad and st["korak"] > 1:
                self.btn(dno, "‹  Back", prethodni).pack(side="right", padx=8)
            return b

        def sledeci():
            st["korak"] += 1
            nacrtaj()

        def prethodni():
            st["korak"] -= 1
            nacrtaj()

        def k1():
            self.napomena(sadrzaj, wraplength=520, tekst="This action deletes ALL animals from the records. After that the animal list, mating plan, "
                                   "relationships and overview will be empty.")
            self.lbl(sadrzaj, "Currently in the records:", 10, "bold").pack(anchor="w", pady=(4, 6))
            for t in (f"•  {n_grla} animals (including deleted and out of herd)",
                      f"•  {n_ist} history records",
                      f"•  table: {self.tab.putanja}"):
                self.lbl(sadrzaj, t, 10, fg=INK2, wraplength=540, justify="left").pack(anchor="w", pady=1)
            self.lbl(sadrzaj, "If you only want to remove some animals, cancel and use normal deletion "
                              "(select the animal and press Delete).", 9, fg=INK2, wraplength=540,
                     justify="left").pack(anchor="w", pady=(14, 0))
            dugmad("I understand, next  ›")

        def k2():
            self.lbl(sadrzaj, "Choose what is deleted:", 10, "bold").pack(anchor="w", pady=(0, 8))
            for val, t, opis in (("sve", "All animals and the whole history",
                                  "The records become completely empty, like a new program."),
                                 ("grla", "Animals only, history is kept",
                                  "History records stay in the file but are not shown until the animals are entered again.")):
                ttk.Radiobutton(sadrzaj, text=t, value=val, variable=st["istorija"],
                                style="Card.TRadiobutton").pack(anchor="w", pady=(6, 0))
                self.lbl(sadrzaj, opis, 9, fg=INK2).pack(anchor="w", padx=26)
            dugmad()

        def k3():
            self.lbl(sadrzaj, "Before deleting, the program makes a backup of all data. Everything can be "
                              "restored from it if you change your mind.", 10, wraplength=540, justify="left").pack(anchor="w")
            if not st["kopija"]:
                try:
                    folder = os.path.join(os.path.dirname(os.path.abspath(self.tab.putanja)), "backups")
                    os.makedirs(folder, exist_ok=True)
                    oznaka = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                    kopije = []
                    for izvor in (self.tab.putanja, self.ist.putanja):
                        if os.path.isfile(izvor):
                            ime, ext = os.path.splitext(os.path.basename(izvor))
                            cilj = os.path.join(folder, f"{ime}_before_delete_{oznaka}{ext}")
                            shutil.copy2(izvor, cilj)
                            kopije.append((izvor, cilj))
                    st["kopija"] = (folder, kopije)
                except OSError as e:
                    self.napomena(sadrzaj, wraplength=520, tekst=f"Backup was not made: {e}\nDeletion has been stopped.")
                    dugmad("Next  ›", stanje="disabled")
                    return
            folder, kopije = st["kopija"]
            self.napomena(sadrzaj, wraplength=520, tekst="Backup has been made.", ok=True)
            for _, cilj in kopije:
                self.lbl(sadrzaj, "•  " + os.path.basename(cilj), 10, fg=INK2).pack(anchor="w", pady=1)
            self.lbl(sadrzaj, f"Folder: {folder}", 9, fg=INK2, wraplength=540, justify="left").pack(
                anchor="w", pady=(8, 0))
            lk = self.lbl(sadrzaj, "Open backup folder", 9, "bold", fg=BRAND, cursor="hand2")
            lk.pack(anchor="w", pady=(6, 0))
            lk.bind("<Button-1>", lambda e: otvori_folder(folder))
            dugmad()

        def k4():
            self.lbl(sadrzaj, "To confirm, type exactly:", 10).pack(anchor="w")
            self.lbl(sadrzaj, "DELETE ALL", 18, "bold", fg=BAD).pack(anchor="w", pady=(6, 10))
            e = ttk.Entry(sadrzaj, textvariable=st["tekst"], width=24, font=self.f(13))
            e.pack(anchor="w")
            info = self.lbl(sadrzaj, "", 9, fg=INK2)
            info.pack(anchor="w", pady=(6, 0))
            b = dugmad(stanje="disabled")

            def provera(*_):
                ok = norm(st["tekst"].get()) == "DELETEALL"
                b.state(["!disabled"] if ok else ["disabled"])
                info.config(text="Confirmed." if ok else "The \"Next\" button turns on when you type the exact text.",
                            fg=OK if ok else INK2)
            tr_id = st["tekst"].trace_add("write", provera)
            e.bind("<Destroy>", lambda ev: st["tekst"].trace_remove("write", tr_id))
            provera()
            e.focus_set()

        def k5():
            if st["gotovo"]:
                folder, kopije = st["kopija"]
                self.napomena(sadrzaj, wraplength=520, tekst="All animals have been deleted.", ok=True)
                self.lbl(sadrzaj, "The backup is saved in the folder:", 10).pack(anchor="w")
                self.lbl(sadrzaj, folder, 9, fg=INK2, wraplength=540, justify="left").pack(anchor="w", pady=(2, 10))
                self.lbl(sadrzaj, "If this was a mistake, click \"Restore data\" – everything will be as before.",
                         9, fg=INK2, wraplength=540, justify="left").pack(anchor="w")
                self.btn(dno, "Restore data", vrati).pack(side="left")
                self.btn(dno, "Close", win.destroy, "Primary").pack(side="right")
                return
            sta = "all animals and the whole history" if st["istorija"].get() == "sve" else "all animals (history is kept)"
            self.lbl(sadrzaj, f"Last step. Clicking the button deletes {sta}.", 11, "bold",
                     wraplength=540, justify="left").pack(anchor="w")
            self.lbl(sadrzaj, f"Number of animals: {n_grla}" + (f"   ·   history records: {n_ist}"
                                                          if st["istorija"].get() == "sve" else ""),
                     10, fg=INK2).pack(anchor="w", pady=(6, 0))
            odb = self.lbl(sadrzaj, "", 10, "bold", fg=WARN)
            odb.pack(anchor="w", pady=(16, 0))
            b = dugmad("Permanently delete all", obrisi_sve, "Danger", stanje="disabled")
            st["odbrojavanje"] = 5

            def tik():
                if not odb.winfo_exists():
                    return
                if st["odbrojavanje"] > 0:
                    odb.config(text=f"The button turns on in {st['odbrojavanje']} s…")
                    st["odbrojavanje"] -= 1
                    win.after(1000, tik)
                else:
                    odb.config(text="You can confirm the deletion.", fg=BAD)
                    b.state(["!disabled"])
            tik()

        def otvori_folder(folder):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(folder)
                elif sys.platform == "darwin":
                    os.system(f'open "{folder}"')
                else:
                    os.system(f'xdg-open "{folder}" &')
            except OSError:
                pass

        def zatvori_prozore():
            for w in (self.karton_win, getattr(self, "ist_win", None), getattr(self, "stablo_win", None)):
                if w is not None and w.winfo_exists():
                    w.destroy()
            self.ui.update(partner="", partner_upit="", partner_vise=[], sA="", sB="")

        def obrisi_sve():
            zatvori_prozore()
            self.tab.redovi = []
            if st["istorija"].get() == "sve":
                self.ist.zapisi = []
            self.promena(f"Deleted all {n_grla} animals.")
            st["gotovo"] = True
            nacrtaj()

        def vrati():
            folder, kopije = st["kopija"]
            try:
                for izvor, cilj in kopije:
                    shutil.copy2(cilj, izvor)
                self.tab.ucitaj()
                self.ist = Istorija(self.tab.putanja)
                self.pods = Podsetnici(self.tab.putanja)
            except Exception as e:
                messagebox.showerror("Restore failed", str(e), parent=win)
                return
            self.model = Model(self.tab.redovi)
            self.prikazi(self.strana)
            self.toast(f"Restored {len(self.tab.redovi)} animals from the backup.")
            win.destroy()

        win.bind("<Escape>", lambda e: win.destroy())
        nacrtaj()
        win.update_idletasks()
        win.geometry(f"+{self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2}"
                     f"+{self.root.winfo_rooty() + max(20, (self.root.winfo_height() - win.winfo_height()) // 3)}")

    def obrisi_grla(self, ids, parent=None):
        ids = [i for i in ids if self.red_za(i) and self.red_za(i).get("status") != "obrisano"]
        if not ids:
            return False
        naziv = ids[0] if len(ids) == 1 else f"{len(ids)} animals ({', '.join(ids[:5])}{'…' if len(ids) > 5 else ''})"
        if not messagebox.askyesno(
                "Delete animal", f"Delete {naziv}?\n\n"
                "The animal disappears from the list and all searches, but stays in the pedigree, "
                "so that the relationships of its offspring stay correct.\n\n"
                "You can restore it on the \"Deleted animals\" page.", parent=parent or self.root):
            return False
        for i in ids:
            self.red_za(i)["status"] = "obrisano"
            self.ist.dodaj(i, AUTO, "Animal deleted from the list (kept in the pedigree)", automatski=True)
        if self.karton_win and self.karton_win.winfo_exists() and getattr(self.karton_win, "grlo", None) in ids:
            self.karton_win.destroy()
        if self.ui.get("partner") in ids:
            self.ui.update(partner="", partner_upit="")
        self.promena(f"Deleted: {', '.join(ids[:3])}{'…' if len(ids) > 3 else ''}")
        return True

    def vrati_grla(self, ids):
        for i in ids:
            r = self.red_za(i)
            if r:
                r["status"] = "aktivno"
                self.ist.dodaj(i, AUTO, "Animal restored to the herd", automatski=True)
        self.promena(f"Restored to herd: {', '.join(ids[:3])}{'…' if len(ids) > 3 else ''}")

    def trajno_obrisi(self, ids):
        m = self.model
        pot = sorted({k for k, pr in m.P.items() for i in ids if i in pr and k not in ids}, key=sort_key)
        upoz = (f"\n\nWarning: {', '.join(pot[:6])}{'…' if len(pot) > 6 else ''} has/have this animal as a parent. "
                "Their ancestors through this animal will be lost and the relationships will be wrong."
                if pot else "")
        if not messagebox.askyesno("Permanent deletion", f"Permanently delete {', '.join(ids[:5])}"
                                   f"{'…' if len(ids) > 5 else ''}? This cannot be undone.{upoz}",
                                   icon="warning"):
            return
        self.tab.redovi = [r for r in self.tab.redovi if r["id"] not in ids]
        self.promena(f"Permanently deleted: {len(ids)} animals")

    def meni_grla(self, tv, ev):
        iid = tv.identify_row(ev.y)
        if not iid or iid not in self.model.A:
            return
        if iid not in tv.selection():
            tv.selection_set(iid)
        sel = [i for i in tv.selection() if i in self.model.A]
        mn = tk.Menu(self.root, tearoff=0, font=self.f(10), bg=SURF, fg=INK, activebackground=SURF2,
                     activeforeground=INK, bd=0)
        if len(sel) == 1:
            mn.add_command(label="Open card", command=lambda: self.karton(iid))
            mn.add_command(label="Edit details", command=lambda: self.forma(iid))
            mn.add_command(label="Change ear tag", command=lambda: self.promena_markice(iid))
            mn.add_command(label="Partner search", command=lambda: self.idi_na_pretragu(iid))
            mn.add_command(label="Pedigree (family tree)", command=lambda: self.stablo(iid))
            mn.add_command(label="Animal history (details and entry)", command=lambda: self.istorija_prozor(iid))
            mn.add_command(label="Add event…", command=lambda: self.dogadjaj_forma(iid))
            mn.add_command(label="Add reminder…", command=lambda: self.podsetnik_forma(iid))
            mn.add_separator()
        mn.add_command(label=f"Move to another herd{'' if len(sel) == 1 else f' ({len(sel)} animals)'}…",
                       command=lambda: self.promena_stada(sel))
        st_ = self.model.A[iid]["stado"]
        if st_:
            mn.add_command(label=f"Rename herd \"{st_}\"…", command=lambda: self.preimenuj_stado(st_))
        mn.add_separator()
        mn.add_command(label=f"Delete{'' if len(sel) == 1 else f' ({len(sel)} animals)'}",
                       foreground=BAD, command=lambda: self.obrisi_grla(sel))
        try:
            mn.tk_popup(ev.x_root, ev.y_root)
        finally:
            mn.grab_release()

    def s_obrisana(self):
        body = self.strana_okvir(scroll=False)
        m = self.model
        obr = m.obrisana()
        self.zaglavlje(body, "Deleted animals",
                       "These animals are not shown in the list or searches, but stay in the pedigree so that the relationships "
                       "of their offspring stay correct.",
                       [("← Back to animals", lambda: self.prikazi("grla"), "Secondary")])
        if not obr:
            self.napomena(body, "There are no deleted animals.", ok=True)
            return
        tb = tk.Frame(body, bg=BG)
        tb.pack(fill="x", pady=(0, 12))
        wrap, tv = self.tabela(body, [("id", "Animal", 105), ("markica", "Ear tag", 125), ("pol", "Sex", 70),
                                      ("rasa", "Breed", 130), ("stado", "Herd", 90),
                                      ("pot", "Offspring in pedigree", 150, "e")], visina=14)
        for i, a in enumerate(obr):
            n = sum(1 for pr in m.P.values() if a["id"] in pr)
            tv.insert("", "end", iid=a["id"], text="  " + a["id"], image=self.img_tag_siv,
                      values=(a["markica"] or "–", m.pol_naziv(a["id"]), a["rasa"] or "–", a["stado"] or "–", n),
                      tags=["odd"] if i % 2 else [])
        izbor = lambda: [i for i in tv.selection() if i in m.A]
        def vrati():
            s_ = izbor()
            if s_:
                self.vrati_grla(s_)
            else:
                self.toast("First select an animal in the table.")
        def trajno():
            s_ = izbor()
            if s_:
                self.trajno_obrisi(s_)
            else:
                self.toast("First select an animal in the table.")
        self.btn(tb, "Restore to herd", vrati, "Primary").pack(side="left")
        self.btn(tb, "Delete permanently", trajno, "Danger").pack(side="left", padx=8)
        self.lbl(tb, "Select an animal (or several with Ctrl), then choose an action.", 9, fg=INK2).pack(side="left", padx=8)
        wrap.pack(fill="both", expand=True)
        tv.bind("<Double-1>", lambda e: self.otvori_iz(tv))

    # ---------- Parenje
    def s_parenje(self):
        body = self.strana_okvir()
        self.zaglavlje(body, "Mating", "The program chooses partners so that the offspring have as little inbreeding as possible. "
                                            f"The risk limit is {pct(self.prag)} (changed on the Overview page).")
        segw = tk.Frame(body, bg=BG)
        segw.pack(fill="x", pady=(0, 16))
        seg = tk.Frame(segw, bg=SURF2)
        seg.pack(side="left")
        for k, t in (("partner", "Partner search"), ("plan", "Plan for the whole herd")):
            akt = self.ui["par_tab"] == k
            b = tk.Label(seg, text=t, font=self.f(10, "bold"), bg=SURF if akt else SURF2, fg=INK if akt else INK2,
                         padx=14, pady=7, cursor="hand2")
            b.pack(side="left", padx=3, pady=3)
            b.bind("<Button-1>", lambda e, k=k: (self.ui.update(par_tab=k), self.prikazi("parenje")))
        if not self.model.grla():
            self.prazno(body)
            return
        (self.partner_view if self.ui["par_tab"] == "partner" else self.plan_view)(body)

    def opis_grla(self, a):
        return f"{a['id']}" + (f"  ·  {a['markica']}" if a["markica"] else "") + f"  ({self.model.pol_naziv(a['id'])})"

    def partner_view(self, body):
        m = self.model
        akt = m.grla(True)
        o, c = self.card(body)
        o.pack(fill="x")
        self.lbl(c, "Partner search", 13, "bold").pack(anchor="w")
        self.lbl(c, "Type the tag or ear tag of an animal and press Enter. The program finds partners of the opposite sex, "
                    "with the least related ones at the top.", 9, fg=INK2, wraplength=760,
                 justify="left").pack(anchor="w", pady=(2, 10))
        row = tk.Frame(c, bg=SURF)
        row.pack(fill="x")
        var = tk.StringVar(value=self.ui.get("partner_upit", self.ui["partner"]))
        svi_opisi = [self.opis_grla(g) for g in akt]
        e = ttk.Combobox(row, textvariable=var, values=svi_opisi, width=34, font=self.f(11))
        e.pack(side="left")
        ph = self.lbl(row, "e.g. K12 or UK 123456 700123", 10, fg="#98A39B", bg=SURF)

        def ph_show(*_):
            if var.get():
                ph.place_forget()
            else:
                ph.place(in_=e, x=10, rely=0.5, anchor="w")
        ph.bind("<Button-1>", lambda ev: e.focus_set())
        ph_show()

        def suzi(ev=None):
            # dok kucas, padajuci spisak prikazuje samo grla koja odgovaraju
            if ev is not None and ev.keysym in ("Return", "KP_Enter", "Down", "Up", "Escape"):
                return
            q = norm(var.get())
            e.configure(values=[d for g, d in zip(akt, svi_opisi)
                                if not q or q in norm(g["id"]) or q in norm(g["markica"])
                                or any(q in norm(x) for x in g["stare"])])
            ph_show()

        def trazi(ev=None):
            upit = var.get().strip()
            if "  ·  " in upit or "  (" in upit:          # izabrano sa spiska
                upit = upit.split("  ")[0]
            self.ui["partner_upit"] = upit
            self.ui["partner"] = ""
            self.ui["partner_vise"] = []
            self.ui["partner_obrisano"] = ""
            if upit:
                x = m.nadji(upit)
                if x and m.A[x].get("obrisano"):
                    self.ui["partner_obrisano"] = x
                elif x and not m.A[x]["osnivac"]:
                    self.ui["partner"] = x
                    self.ui["partner_upit"] = x
                else:
                    q = norm(upit)
                    pogodci = [g["id"] for g in m.grla() if q in norm(g["id"]) or q in norm(g["markica"])
                               or any(q in norm(x) for x in g["stare"])]
                    if len(pogodci) == 1:
                        self.ui["partner"] = self.ui["partner_upit"] = pogodci[0]
                    else:
                        self.ui["partner_vise"] = pogodci[:30]
            self.prikazi("parenje")

        e.bind("<KeyRelease>", suzi)
        e.bind("<Return>", trazi)
        e.bind("<KP_Enter>", trazi)
        e.bind("<<ComboboxSelected>>", trazi)
        self.btn(row, "Find partners", trazi, "Primary").pack(side="left", padx=10)
        row2 = tk.Frame(c, bg=SURF)
        row2.pack(fill="x", pady=(10, 0))
        rv = tk.BooleanVar(value=self.ui["ista_rasa"])
        ttk.Checkbutton(row2, text="Same breed only", variable=rv, style="Card.TCheckbutton",
                        command=lambda: (self.ui.update(ista_rasa=rv.get()), self.prikazi("parenje"))).pack(
            side="left")
        uv = tk.BooleanVar(value=self.ui.get("partner_svi", False))
        ttk.Checkbutton(row2, text="Include animals out of herd", variable=uv, style="Card.TCheckbutton",
                        command=lambda: (self.ui.update(partner_svi=uv.get()), self.prikazi("parenje"))).pack(
            side="left", padx=18)
        if not self.ui["partner"]:
            e.focus_set()

        x = self.ui["partner"]
        upit = self.ui.get("partner_upit", "")
        if not x:
            vise = self.ui.get("partner_vise", [])
            if upit and vise:
                tk.Frame(c, bg=SURF, height=12).pack()
                self.lbl(c, f"Several animals match \"{upit}\". Choose one:", 10, fg=INK2).pack(anchor="w")
                ch = tk.Frame(c, bg=SURF)
                ch.pack(anchor="w", pady=(6, 0))
                for i, g in enumerate(vise):
                    if i and i % 8 == 0:
                        ch = tk.Frame(c, bg=SURF)
                        ch.pack(anchor="w", pady=(4, 0))
                    self.cip(ch, g, lambda g=g: (self.ui.update(partner=g, partner_upit=g, partner_vise=[]),
                                                  self.prikazi("parenje"))).pack(side="left", padx=(0, 6))
            elif upit and self.ui.get("partner_obrisano"):
                tk.Frame(c, bg=SURF, height=12).pack()
                self.napomena(c, f"Animal {self.ui['partner_obrisano']} is deleted. You can restore it on the page "
                                 "Animals → Deleted animals.")
            elif upit:
                tk.Frame(c, bg=SURF, height=12).pack()
                self.napomena(c, f"Animal \"{upit}\" was not found by tag or ear tag.")
            return
        if x not in m.A or m.A[x]["osnivac"]:
            return
        a = m.A[x]
        if not a["pol"]:
            tk.Frame(body, bg=BG, height=16).pack()
            self.napomena(body, f"Animal {x} has no sex entered. Enter the sex in the animal card and try again.")
            return

        # izabrano grlo
        tk.Frame(body, bg=BG, height=16).pack()
        o, c2 = self.card(body, 20, 14)
        o.pack(fill="x")
        top = tk.Frame(c2, bg=SURF)
        top.pack(fill="x")
        self.cip(top, x, lambda: self.karton(x), big=True).pack(side="left")
        t = tk.Frame(top, bg=SURF)
        t.pack(side="left", padx=14)
        self.lbl(t, f"{m.pol_naziv(x)}{', ' + a['rasa'] if a['rasa'] else ''}", 12, "bold").pack(anchor="w")
        o_, mj = m.P[x]
        self.lbl(t, f"ear tag {a['markica'] or '–'}   ·   sire {o_ or '–'}   ·   dam {mj or '–'}"
                    f"   ·   inbreeding {pct(m.inbr(x))}", 9, fg=INK2).pack(anchor="w")

        izvor = m.grla(not self.ui.get("partner_svi", False))
        kand = [(m.kinship(x, k["id"]), m.r(x, k["id"]), k) for k in izvor
                if k["id"] != x and k["vrsta"] == a["vrsta"] and k["pol"] and k["pol"] != a["pol"]
                and (not self.ui["ista_rasa"] or k["rasa"] == a["rasa"])]
        kand.sort(key=lambda t: (t[0], t[1], sort_key(t[2]["id"])))   # najmanje srodni na vrhu
        tk.Frame(body, bg=BG, height=16).pack()
        suprotni = "female" if a["pol"] == "m" else "male"
        if not kand:
            self.napomena(body, f"There are no {suprotni} animals of the same species" +
                          (" and breed" if self.ui["ista_rasa"] else "") + " in the herd.")
            return
        nesrodni = sum(1 for f, r, k in kand if f == 0)
        dobri = sum(1 for f, r, k in kand if f <= self.prag)
        if kand[0][0] > self.prag:
            self.napomena(body, f"No partner in the herd is below the limit of {pct(self.prag)}. "
                                "Consider buying an unrelated animal or semen.")
        else:
            self.napomena(body, f"Found {len(kand)} {suprotni} animals: {nesrodni} completely unrelated, "
                                f"{dobri} below the limit of {pct(self.prag)}. The best choice is at the top.", ok=True)
        wrap, tv = self.tabela(body, [("id", "Partner", 100), ("mk", "Ear tag", 115), ("rasa", "Breed", 90),
                                      ("st", "Age", 80), ("odnos", "Relation", 130), ("r", "Shared genes", 125, "e"),
                                      ("f", "Offspring F", 125, "e"), ("pr", "Rating", 150)],
                               visina=min(len(kand), 14))
        wrap.pack(fill="x")
        for i, (f, r, k) in enumerate(kand):
            t, boja = procena(f, self.prag)
            tv.insert("", "end", iid=k["id"], text=f"  {i + 1}.  {k['id']}", image=self.img_tag,
                      values=(k["markica"] or "–", k["rasa"] or "–", fmt_starost(starost_meseci(k["datum"])) or "–",
                              m.odnos(x, k["id"]), pct(r), pct(f), t),
                      tags=(["odd"] if i % 2 else []) + [{OK: "", WARN: "warn", BAD: "bad"}[boja]]
                           + ([] if k["aktivno"] else ["off"]))
        tv.bind("<Double-1>", lambda e: self.otvori_iz(tv))
        self.lbl(body, "Double-click a partner to open its card.  Shared genes = the share of genes two animals "
                       "have from the same ancestors (parent and offspring or full sibling ≈ 50%, half sibling ≈ 25%, "
                       "unrelated 0%).", 9, fg=INK2, wraplength=900, justify="left").pack(anchor="w", pady=(8, 0))

    def napomena(self, parent, tekst, ok=False, wraplength=820):
        f = tk.Frame(parent, bg=OK_BG if ok else WARN_BG)
        f.pack(fill="x", pady=(0, 14))
        self.lbl(f, tekst, 10, fg=OK if ok else WARN, justify="left", wraplength=wraplength).pack(
            anchor="w", padx=14, pady=10)

    def plan_view(self, body):
        m = self.model
        akt = m.grla(True)
        o, c = self.card(body)
        o.pack(fill="x")
        row = tk.Frame(c, bg=SURF)
        row.pack(fill="x")
        self.lbl(row, "Max. females per male:", 10, "bold", fg=INK2).pack(side="left")
        mv = tk.StringVar(value=str(self.ui["plan_max"] or ""))
        sp = ttk.Spinbox(row, from_=0, to=999, textvariable=mv, width=6, font=self.f(10))
        sp.pack(side="left", padx=8)
        self.lbl(row, "(empty = no limit)", 9, fg=INK2).pack(side="left")
        rv = tk.BooleanVar(value=self.ui["ista_rasa"])
        ttk.Checkbutton(row, text="Same breed only", variable=rv, style="Card.TCheckbutton",
                        command=lambda: (self.ui.update(ista_rasa=rv.get()), self.prikazi("parenje"))).pack(
            side="left", padx=18)

        def prim(*_):
            try:
                self.ui["plan_max"] = max(0, int(mv.get() or 0))
            except ValueError:
                self.ui["plan_max"] = 0
            self.prikazi("parenje")
        sp.config(command=prim)
        sp.bind("<Return>", prim)
        sp.bind("<FocusOut>", lambda e: str(self.ui["plan_max"] or "") != mv.get() and prim())
        tk.Frame(body, bg=BG, height=16).pack()

        zenke = [a for a in akt if a["pol"] == "z"]
        muz = [a for a in akt if a["pol"] == "m"]
        if not zenke or not muz:
            self.napomena(body, f"The plan needs both female and male animals in the herd "
                                f"(now: {len(zenke)} female, {len(muz)} male).")
            return
        mx = self.ui["plan_max"]
        opc = {z["id"]: sorted((m.kinship(z["id"], x["id"]), x["id"]) for x in muz
                               if x["vrsta"] == z["vrsta"] and (not self.ui["ista_rasa"] or x["rasa"] == z["rasa"]))
               for z in zenke}
        kor = {x["id"]: 0 for x in muz}
        plan = []
        def tez(z):
            l = opc[z["id"]]
            return (sum(1 for f, _ in l if f <= self.prag), l[0][0] if l else 1)
        for z in sorted(zenke, key=tez):
            izbor = None
            for f, x in opc[z["id"]]:
                if mx and kor[x] >= mx:
                    continue
                kand = [(kor[x2], sort_key(x2), x2) for f2, x2 in opc[z["id"]]
                        if abs(f2 - f) < 1e-12 and not (mx and kor[x2] >= mx)]
                izbor = (f, min(kand)[2])
                break
            if izbor:
                kor[izbor[1]] += 1
            plan.append((z["id"], izbor[1] if izbor else None, izbor[0] if izbor else None))
        plan.sort(key=lambda t: sort_key(t[0]))
        fs = [f for _, x, f in plan if x]
        prob = sum(1 for _, x, f in plan if not x or f > self.prag)

        st = tk.Frame(body, bg=BG)
        st.pack(fill="x", pady=(0, 14))
        for i, (b, t, boja) in enumerate(((str(len(zenke)), "female animals", INK), (str(len(muz)), "male animals", INK),
                                          (pct(sum(fs) / len(fs)) if fs else "–", "average inbreeding of offspring", INK),
                                          (str(prob), "matings above the limit", BAD if prob else OK))):
            self.stat(st, b, t, boja).grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
            st.columnconfigure(i, weight=1, uniform="s")
        if prob:
            self.napomena(body, f"{prob} matings are above the limit or without a partner. "
                                "Consider buying a new unrelated male or semen.")
        wrap, tv = self.tabela(body, [("z", "Female", 120), ("rasa", "Breed", 120), ("stado", "Herd", 90),
                                      ("x", "Partner", 110), ("rx", "Partner breed", 120),
                                      ("f", "Offspring inbreeding", 150, "e"), ("pr", "Rating", 150)],
                               visina=min(len(plan), 14))
        wrap.pack(fill="x")
        for i, (z, x, f) in enumerate(plan):
            a = m.A[z]
            if x:
                t, boja = procena(f, self.prag)
                tag = {OK: "", WARN: "warn", BAD: "bad"}[boja]
            else:
                t, tag = "no free partner", "bad"
            tv.insert("", "end", iid=z, text="  " + z, image=self.img_tag,
                      values=(a["rasa"] or "–", a["stado"] or "–", x or "–", m.A[x]["rasa"] if x else "",
                              pct(f) if x else "", t), tags=(["odd"] if i % 2 else []) + [tag])
        tv.bind("<Double-1>", lambda e: self.otvori_iz(tv))

        tk.Frame(body, bg=BG, height=16).pack()
        o, c = self.card(body)
        o.pack(fill="x")
        self.lbl(c, "Load on males", 13, "bold").pack(anchor="w")
        self.lbl(c, "How many females are assigned to each male in this plan.", 9, fg=INK2).pack(anchor="w", pady=(0, 10))
        for x in sorted(muz, key=lambda a: -kor[a["id"]]):
            r = tk.Frame(c, bg=SURF)
            r.pack(fill="x", pady=4)
            self.cip(r, x["id"], lambda i=x["id"]: self.karton(i)).pack(side="left")
            self.lbl(r, x["rasa"], 9, fg=INK2, width=14, anchor="w").pack(side="left", padx=8)
            n = kor[x["id"]]
            self.lbl(r, f"{n} " + ("female" if n == 1 else "females"), 10, "bold",
                     width=10, anchor="e").pack(side="right")
            pb = ttk.Progressbar(r, style="Green.Horizontal.TProgressbar", maximum=len(zenke), value=n)
            pb.pack(side="left", fill="x", expand=True, padx=8)

    # ---------- Srodstvo
    def s_srodstvo(self):
        body = self.strana_okvir()
        self.zaglavlje(body, "Relationship", "Choose two animals and see how related they are and what their offspring would be like.")
        m = self.model
        sva = m.grla()
        if not sva:
            self.prazno(body)
            return
        opisi = [self.opis_grla(a) for a in sva]
        o, c = self.card(body)
        o.pack(fill="x")
        row = tk.Frame(c, bg=SURF)
        row.pack(fill="x")
        vars_ = {}
        for i, (k, t) in enumerate((("sA", "First animal"), ("sB", "Second animal"))):
            f = tk.Frame(row, bg=SURF)
            f.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 16, 0))
            row.columnconfigure(i, weight=1)
            self.lbl(f, t, 10, "bold", fg=INK2).pack(anchor="w")
            v = tk.StringVar(value=next((d for a, d in zip(sva, opisi) if a["id"] == self.ui[k]), ""))
            cb = ttk.Combobox(f, textvariable=v, values=opisi, state="readonly", font=self.f(10))
            cb.pack(fill="x", pady=(6, 0))
            cb.bind("<<ComboboxSelected>>",
                    lambda e, k=k, v=v: (self.ui.update({k: sva[opisi.index(v.get())]["id"]}), self.prikazi("srodstvo")))
            vars_[k] = v
        a, b = self.ui["sA"], self.ui["sB"]
        if a in m.A and b in m.A and a != b:
            A, B = m.A[a], m.A[b]
            if A["vrsta"] != B["vrsta"]:
                tk.Frame(c, bg=SURF, height=14).pack()
                self.napomena(c, f"{a} is {A['vrsta']}, and {b} is {B['vrsta']} – different species.")
            else:
                f, r = m.kinship(a, b), m.r(a, b)
                v = tk.Frame(c, bg=SURF)
                v.pack(fill="x", pady=(18, 0))
                t, boja = procena(f, self.prag)
                for i, (big, small, col) in enumerate(((pct(r), f"shared genes (r = {r:.3f})", INK),
                                                       (m.odnos(a, b), "odnos", INK),
                                                       (pct(f), f"inbreeding of the {nazivi(A['vrsta'])[2]}", boja))):
                    fr = tk.Frame(v, bg=SURF2)
                    fr.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
                    v.columnconfigure(i, weight=1, uniform="v")
                    self.lbl(fr, big, 22 if i != 1 else 15, "bold", fg=col).pack(anchor="w", padx=16, pady=(12, 0))
                    self.lbl(fr, small, 9, fg=INK2).pack(anchor="w", padx=16, pady=(0, 12 if i != 2 else 4))
                    if i == 2:
                        self.badge(fr, t, boja).pack(anchor="w", padx=16, pady=(0, 12))
                self.lbl(c, "Common ancestors", 10, "bold").pack(anchor="w", pady=(16, 6))
                zaj = sorted((m.preci(a) | {a}) & (m.preci(b) | {b}), key=sort_key)
                ch = tk.Frame(c, bg=SURF)
                ch.pack(anchor="w")
                if zaj:
                    for z in zaj:
                        self.cip(ch, z, lambda z=z: self.karton(z), siv=m.A[z]["osnivac"]).pack(side="left", padx=(0, 6))
                else:
                    self.lbl(ch, "no known common ancestors", 10, fg=INK2).pack(side="left")

        tk.Frame(body, bg=BG, height=16).pack()
        o, c = self.card(body)
        o.pack(fill="x")
        self.lbl(c, "Most related pairs in the herd", 13, "bold").pack(anchor="w")
        self.lbl(c, "Animals that should be kept apart when mating.", 9, fg=INK2).pack(anchor="w", pady=(0, 10))
        akt = m.grla(True)
        min_r = 0.25 if len(akt) > 300 else 1e-9
        po_v = {}
        for x in akt:
            po_v.setdefault(x["vrsta"], []).append(x["id"])
        parovi = []
        for ids in po_v.values():
            for i in range(len(ids)):
                for j in range(i + 1, len(ids)):
                    r = m.r(ids[i], ids[j])
                    if r >= min_r:
                        parovi.append((r, ids[i], ids[j]))
        parovi.sort(reverse=True)
        if not parovi:
            self.lbl(c, "No related pairs according to the recorded pedigree.", 10, fg=INK2).pack(anchor="w")
            return
        wrap, tv = self.tabela(c, [("a", "Animal A", 120), ("b", "Animal B", 120), ("od", "Relation", 200),
                                   ("r", "Shared genes", 120, "e")], visina=min(len(parovi), 12))
        wrap.pack(fill="x")
        for i, (r, x, y) in enumerate(parovi[:40]):
            tv.insert("", "end", text="  " + x, image=self.img_tag,
                      values=(y, m.odnos(x, y), pct(r)), tags=["odd"] if i % 2 else [])

    # ---------- Pregled
    def s_pregled(self):
        body = self.strana_okvir()
        h = self.zaglavlje(body, "Herd overview", "Inbreeding status, most common sires and data errors.")
        pf = tk.Frame(h, bg=BG)
        pf.pack(side="right", anchor="s")
        self.lbl(pf, "Risk limit", 9, "bold", fg=INK2).pack(anchor="w")
        opcije = {"3.125% – strict": 0.03125, "6.25% – recommended": 0.0625, "12.5% – mild": 0.125}
        pv = tk.StringVar(value=next((k for k, v in opcije.items() if abs(v - self.prag) < 1e-9), "6.25% – recommended"))
        cb = ttk.Combobox(pf, textvariable=pv, values=list(opcije), state="readonly", width=22, font=self.f(10))
        cb.pack()
        cb.bind("<<ComboboxSelected>>", lambda e: (setattr(self, "prag", opcije[pv.get()]),
                                                   self.sacuvaj_podesavanja(), self.prikazi("pregled")))
        m = self.model
        sva = m.grla()
        if not sva:
            self.prazno(body)
            return
        akt = [a for a in sva if a["aktivno"]]
        F = [m.inbr(a["id"]) for a in akt]
        iznad = sum(f > self.prag for f in F)
        st = tk.Frame(body, bg=BG)
        st.pack(fill="x", pady=(0, 16))
        for i, (b, t, col) in enumerate(((str(len(akt)), "animals in herd", INK),
                                         (str(sum(a["pol"] == "z" for a in akt)), "female", INK),
                                         (str(sum(a["pol"] == "m" for a in akt)), "male", INK),
                                         (pct(sum(F) / len(F)) if F else "–", "average inbreeding", INK),
                                         (str(iznad), "above the limit", BAD if iznad else OK))):
            self.stat(st, b, t, col).grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
            st.columnconfigure(i, weight=1, uniform="s")

        dva = tk.Frame(body, bg=BG)
        dva.pack(fill="x")
        dva.columnconfigure(0, weight=1, uniform="d")
        dva.columnconfigure(1, weight=1, uniform="d")
        o, c = self.card(dva)
        o.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.lbl(c, "Most common sires", 13, "bold").pack(anchor="w")
        self.lbl(c, "When one male has too many offspring, the next generation becomes related.", 9, fg=INK2,
                 wraplength=380, justify="left").pack(anchor="w", pady=(0, 10))
        oc = {}
        for a in akt:
            o_ = m.P[a["id"]][0]
            if o_:
                oc[o_] = oc.get(o_, 0) + 1
        ocl = sorted(oc.items(), key=lambda t: -t[1])[:8]
        if ocl and ocl[0][1] / len(akt) > 0.3:
            self.napomena(c, "One sire has more than 30% of the offspring in the herd.")
        if not ocl:
            self.lbl(c, "No sires entered yet.", 10, fg=INK2).pack(anchor="w")
        for x, n in ocl:
            r = tk.Frame(c, bg=SURF)
            r.pack(fill="x", pady=4)
            self.cip(r, x, lambda x=x: self.karton(x), siv=m.A[x]["osnivac"]).pack(side="left")
            self.lbl(r, f"{n} ({round(n / len(akt) * 100)}%)", 10, "bold", width=10, anchor="e").pack(side="right")
            ttk.Progressbar(r, style="Green.Horizontal.TProgressbar", maximum=len(akt), value=n).pack(
                side="left", fill="x", expand=True, padx=10)

        o, c = self.card(dva)
        o.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self.lbl(c, "Animals with the highest inbreeding", 13, "bold").pack(anchor="w")
        self.lbl(c, "These animals come from related parents.", 9, fg=INK2).pack(anchor="w", pady=(0, 10))
        inb = sorted(((f, a) for f, a in zip(F, akt) if f > 0), key=lambda t: -t[0])[:10]
        if not inb:
            self.napomena(c, "No animal in the herd is inbred.", ok=True)
        for f, a in inb:
            r = tk.Frame(c, bg=SURF)
            r.pack(fill="x", pady=4)
            self.cip(r, a["id"], lambda x=a["id"]: self.karton(x)).pack(side="left")
            o_, mj = m.P[a["id"]]
            self.lbl(r, f"{o_} × {mj}", 9, fg=INK2).pack(side="left", padx=10)
            self.badge(r, pct(f), BAD if f > self.prag else WARN if f > self.prag / 2 else OK).pack(side="right")

        tk.Frame(body, bg=BG, height=16).pack()
        grupe = {}
        for a in akt:
            grupe.setdefault((a["vrsta"], a["rasa"] or "–", a["stado"] or "–"), []).append(a)
        wrap, tv = self.tabela(body, [("v", "Species", 100), ("r", "Breed", 130), ("s", "Herd", 100),
                                      ("n", "Animals", 70, "e"), ("m", "Males", 70, "e"), ("z", "Females", 70, "e"),
                                      ("pf", "Average F", 100, "e"), ("mf", "Highest F", 100, "e")],
                               visina=min(len(grupe), 8), tree=False)
        wrap.pack(fill="x")
        for i, ((v, r, s), g) in enumerate(sorted(grupe.items(), key=lambda t: -len(t[1]))):
            f = [m.inbr(a["id"]) for a in g]
            tv.insert("", "end", values=(v, r, s, len(g), sum(a["pol"] == "m" for a in g),
                                         sum(a["pol"] == "z" for a in g), pct(sum(f) / len(f)), pct(max(f))),
                      tags=["odd"] if i % 2 else [])

        tk.Frame(body, bg=BG, height=16).pack()
        o, c = self.card(body)
        o.pack(fill="x")
        self.lbl(c, "Data check", 13, "bold").pack(anchor="w", pady=(0, 8))
        if not m.warn:
            self.napomena(c, "No errors found in the pedigree.", ok=True)
        for w in m.warn[:40]:
            r = tk.Frame(c, bg=SURF)
            r.pack(fill="x", pady=3)
            self.badge(r, "!", WARN).pack(side="left", anchor="n")
            self.lbl(r, w, 10, justify="left", wraplength=780).pack(side="left", padx=8, anchor="w")

        tk.Frame(body, bg=BG, height=16).pack()
        self.kartica_plodnosti(body)
        tk.Frame(body, bg=BG, height=16).pack()
        self.kartica_kopija(body)

        # opasna zona
        tk.Frame(body, bg=BG, height=16).pack()
        o = tk.Frame(body, bg=BAD)
        o.pack(fill="x")
        c = tk.Frame(o, bg=SURF)
        c.pack(fill="both", expand=True, padx=1, pady=1)
        red_ = tk.Frame(c, bg=SURF)
        red_.pack(fill="x", padx=20, pady=16)
        lev = tk.Frame(red_, bg=SURF)
        lev.pack(side="left", fill="x", expand=True)
        self.lbl(lev, "Danger zone", 13, "bold", fg=BAD).pack(anchor="w")
        self.lbl(lev, "Delete all animals from the records. Done in 5 steps, with an automatic backup.",
                 9, fg=INK2).pack(anchor="w")
        self.btn(red_, "Delete all animals…", self.brisanje_svih, "Danger").pack(side="right")

    # ---------- karton
    def karton(self, x):
        m = self.model
        if x not in m.A:
            return
        if self.karton_win and self.karton_win.winfo_exists():
            win = self.karton_win
            for w in win.winfo_children():
                w.destroy()
        else:
            win = tk.Toplevel(self.root, bg=SURF)
            win.geometry(f"480x{min(760, self.root.winfo_height())}+"
                         f"{self.root.winfo_rootx() + max(0, self.root.winfo_width() - 500)}+{self.root.winfo_rooty() + 20}")
            win.transient(self.root)
            win.bind("<Escape>", lambda e: win.destroy())
            self.karton_win = win
        win.grlo = x
        a = m.A[x]
        win.title(f"Animal card {x}")
        sf = ScrollFrame(win, bg=SURF)
        sf.pack(fill="both", expand=True)
        b = tk.Frame(sf.inner, bg=SURF)
        b.pack(fill="both", expand=True, padx=24, pady=22)
        top = tk.Frame(b, bg=SURF)
        top.pack(fill="x")
        self.cip(top, x, big=True, siv=a["osnivac"]).pack(side="left")
        t = tk.Frame(top, bg=SURF)
        t.pack(side="left", padx=14)
        self.lbl(t, m.pol_naziv(x) + (", " + a["rasa"] if a["rasa"] else ""), 12, "bold").pack(anchor="w")
        if a["rasa"].startswith("hybrid"):
            o_h, m_h = m.P.get(x, (None, None))
            ro_ = m.A[o_h]["rasa"] if o_h in m.A else ""
            rm_ = m.A[m_h]["rasa"] if m_h in m.A else ""
            if ro_ or rm_:
                self.lbl(t, f"sire {ro_ or '?'} × dam {rm_ or '?'}", 9, fg=INK2).pack(anchor="w")
        self.lbl(t, "known only as a parent" if a["osnivac"] else "deleted (kept in the pedigree)"
                 if a.get("obrisano") else "in herd" if a["aktivno"] else "not in herd",
                 9, fg=BAD if a.get("obrisano") else INK2).pack(anchor="w")

        def sekcija(naslov):
            self.lbl(b, naslov, 10, "bold").pack(anchor="w", pady=(18, 6))
            fr = tk.Frame(b, bg=SURF)
            fr.pack(anchor="w", fill="x")
            return fr

        if a["osnivac"]:
            self.lbl(b, "This animal is entered only as a parent. Add it as an animal if you want to enter\n"
                        "its details and ancestors.", 10, fg=INK2, justify="left").pack(anchor="w", pady=(14, 8))
            self.btn(b, f"Add {x} as an animal", lambda: (win.destroy(), self.forma(preset=dict(
                id=x, pol=a["pol"], vrsta=a["vrsta"]))), "Primary").pack(anchor="w")
        else:
            kv = tk.Frame(b, bg=SURF)
            kv.pack(fill="x", pady=(16, 0))
            st = starost_meseci(a["datum"])
            mr = tk.Frame(kv, bg=SURF)
            mr.grid(row=0, column=1, sticky="w", padx=16, pady=3)
            self.lbl(mr, a["markica"] or "–", 10, "bold").pack(side="left")
            lnk = self.lbl(mr, "change", 9, "bold", fg=BRAND, cursor="hand2")
            lnk.pack(side="left", padx=10)
            lnk.bind("<Button-1>", lambda e: self.promena_markice(x, parent=win))
            por = a["poreklo"] + (f" – {a['poreklo_od']}" if a["poreklo_od"] else "") if a["poreklo"] else a["poreklo_od"]
            self.lbl(kv, "Origin", 10, fg=INK2).grid(row=6, column=0, sticky="nw", pady=3)
            self.lbl(kv, por or "not entered", 10, "bold" if por else "normal", fg=INK if por else INK2,
                     wraplength=260, justify="left").grid(row=6, column=1, sticky="w", padx=16, pady=3)
            if a["stare"]:
                self.lbl(kv, "Previous ear tags", 10, fg=INK2).grid(row=5, column=0, sticky="nw", pady=3)
                self.lbl(kv, ", ".join(a["stare"]), 10, fg=INK2, wraplength=260, justify="left").grid(
                    row=5, column=1, sticky="w", padx=16, pady=3)
            for i, (k, v) in enumerate((("Ear tag", None), ("Species", a["vrsta"]),
                                        ("Born", (fmt_datum(a["datum"]) + (f"  ({fmt_starost(st)})" if st is not None else ""))
                                         if a["datum"] else "–"),
                                        ("Herd", a["stado"] or "–"))):
                self.lbl(kv, k, 10, fg=INK2).grid(row=i, column=0, sticky="w", pady=3)
                if v is not None and k != "Herd":
                    self.lbl(kv, v, 10, "bold").grid(row=i, column=1, sticky="w", padx=16, pady=3)
            sr = tk.Frame(kv, bg=SURF)
            sr.grid(row=3, column=1, sticky="w", padx=16, pady=3)
            self.lbl(sr, a["stado"] or "–", 10, "bold").pack(side="left")
            lns = self.lbl(sr, "change", 9, "bold", fg=BRAND, cursor="hand2")
            lns.pack(side="left", padx=10)
            lns.bind("<Button-1>", lambda e: self.promena_stada([x], parent=win))
            f = m.inbr(x)
            self.lbl(kv, "Inbreeding", 10, fg=INK2).grid(row=4, column=0, sticky="w", pady=3)
            self.badge(kv, pct(f), BAD if f > self.prag else WARN if f > self.prag / 2 else OK if f else INK2).grid(
                row=4, column=1, sticky="w", padx=16, pady=3)
            kar = self.opis_karence(x)
            if kar:
                self.lbl(kv, "Withdrawal", 10, fg=BAD).grid(row=7, column=0, sticky="nw", pady=3)
                self.lbl(kv, kar, 10, "bold", fg=BAD, wraplength=260, justify="left").grid(
                    row=7, column=1, sticky="w", padx=16, pady=3)

        o, mj = m.P[x]
        fr = sekcija("Parents")
        for p, t in ((o, "sire unknown"), (mj, "dam unknown")):
            if p:
                self.cip(fr, p, lambda p=p: self.karton(p), siv=m.A[p]["osnivac"]).pack(side="left", padx=(0, 6))
            else:
                self.lbl(fr, f" {t} ", 9, fg=INK2, bg=SURF2).pack(side="left", padx=(0, 6))
        dede = [r for p in (o, mj) if p for r in m.P.get(p, ()) if r]
        if dede:
            fr = sekcija("Grandparents")
            for d in dede:
                self.cip(fr, d, lambda d=d: self.karton(d), siv=m.A[d]["osnivac"]).pack(side="left", padx=(0, 6))
        if a["pol"] == "z" and not a["osnivac"]:
            pl = self.plodnost(x)
            if pl["porodjaja"] or pl["parenja"]:
                fr = sekcija("Fertility")
                tab_ = tk.Frame(fr, bg=SURF2)
                tab_.pack(fill="x")
                for i, (k_, v_) in enumerate((
                        ("Births", pl["porodjaja"]),
                        ("Born alive", f"{pl['zivih']}" + (f"  (average {pl['prosek']:.1f} per birth)"
                                                              if pl["porodjaja"] else "")),
                        ("Stillborn", pl["mrtvih"] or "–"),
                        ("Interval between births", f"{round(pl['razmak'])} days" if pl["razmak"] else "–"),
                        ("Matings / failed", f"{pl['parenja']} / {pl['neuspela']}"),
                        ("First birth at", f"{pl['prvi_mes']} months" if pl["prvi_mes"] is not None else "–"))):
                    self.lbl(tab_, k_, 9, fg=INK2, bg=SURF2).grid(row=i, column=0, sticky="w", padx=(12, 10), pady=2)
                    self.lbl(tab_, str(v_), 9, "bold", bg=SURF2).grid(row=i, column=1, sticky="w", padx=(0, 12), pady=2)
        udeli = [(d, p) for d, p in m.udeli_predaka(x) if p not in (o, mj)][:6]
        if udeli:
            fr = sekcija("Share of genes from ancestors")
            for d, p in udeli:
                r = tk.Frame(fr, bg=SURF)
                r.pack(fill="x", pady=2)
                self.cip(r, p, lambda p=p: self.karton(p), siv=m.A[p]["osnivac"]).pack(side="left")
                self.lbl(r, pct(d), 10, "bold", width=7, anchor="e").pack(side="right")
                ttk.Progressbar(r, style="Green.Horizontal.TProgressbar", maximum=100, value=d * 100).pack(
                    side="left", fill="x", expand=True, padx=10)
        if m.putanje(x):
            lk = self.lbl(b, f"Show all ancestors ({len(m.putanje(x))})  ›", 9, "bold", fg=BRAND, cursor="hand2")
            lk.pack(anchor="w", pady=(6, 0))
            lk.bind("<Button-1>", lambda e: self.stablo(x, 4, "svi"))
        n_boc = len(self.bocni_srodnici(x))
        if n_boc:
            lb_ = self.lbl(b, f"Collateral relatives – siblings, cousins ({n_boc})  ›", 9, "bold", fg=BRAND, cursor="hand2")
            lb_.pack(anchor="w", pady=(4, 0))
            lb_.bind("<Button-1>", lambda e: self.stablo(x, 4, "bocni"))
        pot = sorted((k for k, pr in m.P.items() if x in pr), key=sort_key)
        fr = sekcija(f"Offspring ({len(pot)})")
        if not pot:
            self.lbl(fr, "none recorded", 10, fg=INK2).pack(side="left")
        red = None
        for i, p in enumerate(pot[:40]):
            if i % 6 == 0:
                red = tk.Frame(fr, bg=SURF)
                red.pack(anchor="w", pady=2)
            self.cip(red, p, lambda p=p: self.karton(p)).pack(side="left", padx=(0, 6))

        n_svih = sum(len(v) for v in self.potomci_po_kolenima(x).values())
        if n_svih > len(pot):
            lpo = self.lbl(b, f"All descendants – children, grandchildren ({n_svih})  ›", 9, "bold", fg=BRAND, cursor="hand2")
            lpo.pack(anchor="w", pady=(4, 0))
            lpo.bind("<Button-1>", lambda e: self.stablo(x, 4, "potomci"))
        if a["aktivno"] and a["pol"] and not a["osnivac"]:
            kand = sorted((m.kinship(x, k["id"]), k["id"]) for k in m.grla(True)
                          if k["id"] != x and k["vrsta"] == a["vrsta"] and k["pol"] != a["pol"])[:5]
            if kand:
                fr = sekcija("Best partners")
                for f, k in kand:
                    r = tk.Frame(fr, bg=SURF)
                    r.pack(fill="x", pady=3)
                    self.cip(r, k, lambda k=k: self.karton(k)).pack(side="left")
                    self.lbl(r, f"genes {pct(m.r(x, k))}", 9, fg=INK2).pack(side="left", padx=8)
                    t, boja = procena(f, self.prag)
                    self.badge(r, t, boja).pack(side="right")
                    self.lbl(r, pct(f), 10, "bold").pack(side="right", padx=10)

        if not a["osnivac"]:
            obv = [o for o in self.obavestenja(x) if o["grlo"] == x or o.get("podsetnik")]
            fr = sekcija(f"Notifications ({len(obv)})")
            if not obv:
                self.lbl(fr, "none", 10, fg=INK2).pack(anchor="w")
            for o in obv[:5]:
                self.red_obavestenja(fr, o, kompakt=True)
            lp = self.lbl(b, "+ Add reminder", 9, "bold", fg=BRAND, cursor="hand2")
            lp.pack(anchor="w", pady=(4, 0))
            lp.bind("<Button-1>", lambda e: self.podsetnik_forma(x, parent=win))
            zap = self.ist.za(x)
            fr = sekcija(f"History ({len(zap)})")
            sz = tk.Frame(fr, bg=SURF)
            sz.pack(fill="x", pady=(0, 6))
            for i, (naziv, vrste) in enumerate(GRUPE):
                zz = [z for z in zap if z["dogadjaj"] in vrste]
                k = tk.Frame(sz, bg=SURF2, cursor="hand2")
                k.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
                sz.columnconfigure(i, weight=1, uniform="g")
                a1 = tk.Label(k, text=f"{naziv}: {len(zz)}", font=self.f(9, "bold"), bg=SURF2, fg=INK)
                a1.pack(anchor="w", padx=8, pady=(5, 0))
                a2 = tk.Label(k, text=zz[0]["datum"] if zz else "–", font=self.f(8), bg=SURF2, fg=INK2)
                a2.pack(anchor="w", padx=8, pady=(0, 5))
                for w in (k, a1, a2):
                    w.bind("<Button-1>", lambda e, n=naziv: (self.ui.setdefault("ist_" + x, {"vrsta": "All events",
                           "auto": True, "q": ""}).update(grupa=n, vrsta="All events"), self.istorija_prozor(x)))
            if not zap:
                self.lbl(fr, "no records", 10, fg=INK2).pack(anchor="w")
            for z in zap[:5]:
                r = tk.Frame(fr, bg=SURF)
                r.pack(fill="x", pady=2)
                self.lbl(r, z["datum"], 9, "bold", fg=INK2, width=11, anchor="w").pack(side="left", anchor="n")
                tekst = (z["dogadjaj"] if z["dogadjaj"] != AUTO else "") + \
                        (" – " if z["dogadjaj"] != AUTO and (z["opis"] or z["partner"] or z["vrednost"]) else "") + \
                        ", ".join(v for v in (z["opis"], (f"partner {z['partner']}" if z["partner"] else ""),
                                              z["vrednost"]) if v)
                self.lbl(r, tekst, 9, fg=INK2 if z["automatski"] else INK, wraplength=330,
                         justify="left").pack(side="left", anchor="w")
            lr = tk.Frame(b, bg=SURF)
            lr.pack(anchor="w", pady=(6, 0))
            l1 = self.lbl(lr, "+ Add event", 9, "bold", fg=BRAND, cursor="hand2")
            l1.pack(side="left")
            l1.bind("<Button-1>", lambda e: self.dogadjaj_forma(x, parent=win))
            l2 = self.lbl(lr, "Full history  ›", 9, "bold", fg=BRAND, cursor="hand2")
            l2.pack(side="left", padx=16)
            l2.bind("<Button-1>", lambda e: self.istorija_prozor(x))

        act1 = tk.Frame(b, bg=SURF)
        act1.pack(anchor="w", pady=(24, 0))
        if not a["osnivac"]:
            self.btn(act1, "Edit details", lambda: (win.destroy(), self.forma(x)), "Primary").pack(side="left", padx=(0, 8))
        self.btn(act1, "Pedigree – tree", lambda: self.stablo(x)).pack(side="left")
        if not a["osnivac"]:
            act = tk.Frame(b, bg=SURF)
            act.pack(anchor="w", pady=(8, 0))
            dete = dict(vrsta=a["vrsta"], rasa=a["rasa"], stado=a["stado"])
            dete["otac" if a["pol"] == "m" else "majka"] = x
            self.btn(act, "Add offspring", lambda: (win.destroy(), self.forma(preset=dete))).pack(side="left", padx=(0, 8))
            if a.get("obrisano"):
                self.btn(act, "Restore to herd", lambda: self.vrati_grla([x])).pack(side="left")
            else:
                if a["pol"]:
                    self.btn(act, "All partners", lambda: (win.destroy(), self.idi_na_pretragu(x))).pack(side="left")
                self.btn(act, "Delete", lambda: self.obrisi_grla([x], parent=win), "Danger").pack(side="left", padx=8)
        win.lift()

    # ---------- poreklo / rodoslovno stablo
    # ---------- istorija
    def opis_zapisa(self, z):
        dod = []
        if z["partner"]:
            dod.append(f"partner {z['partner']}")
        if z["vrednost"]:
            dod.append(z["vrednost"])
        kar = [f"{n} {z.get(k)} days" for k, n in (("karenca_meso", "meat"), ("karenca_mleko", "milk")) if z.get(k)]
        if kar:
            dod.append("withdrawal: " + ", ".join(kar))
        return " · ".join(v for v in [z["opis"]] + dod if v)

    def tabela_istorije(self, parent, zapisi, sa_grlom, visina=14):
        kol = [("datum", "Date", 100)] + ([("grlo", "Animal", 90)] if sa_grlom else []) + \
              [("dog", "Event", 160), ("opis", "Description", 250), ("ishod", "Young", 185)]
        wrap, tv = self.tabela(parent, kol, visina=visina, tree=False)
        tv.tag_configure("auto", foreground=INK2)
        tv.tag_configure("ceka", foreground=BRAND)
        for i, z in enumerate(zapisi):
            vals = [z["datum"]] + ([z["grlo"]] if sa_grlom else []) + \
                   [z["dogadjaj"] if z["dogadjaj"] != AUTO else "· record", self.opis_zapisa(z),
                    self.opis_ishoda(z)]
            tv.insert("", "end", iid=str(self.ist.zapisi.index(z)), values=vals,
                      tags=(["odd"] if i % 2 else []) + (["auto"] if z["automatski"] else []) +
                           (["ceka"] if self.ceka_porodjaj(z) else []))
        return wrap, tv

    # ---------- parenje -> porođaj -> podmladak
    def roditelji_parenja(self, z):
        """(majka, otac) iz zapisa o parenju, bez obzira kod koga je parenje upisano."""
        m = self.model
        g = z["grlo"]
        p = (m.nadji(z.get("partner")) or z.get("partner") or "").strip() or None
        if g in m.A and m.A[g]["pol"] == "m":
            return p, g
        return g, p

    def profil_gravidnosti(self, vrsta, rasa=""):
        """(ključ, naziv, najmanje, najviše, prosek) dana gravidnosti – po rasi majke, pa po vrsti."""
        v = vrsta_norm(vrsta or "")
        r = norm_rasa(rasa or "")
        kljuc = None
        for k, (naziv, vr, prepoznaj, *_d) in GRAVIDNOST_PROFILI.items():
            if prepoznaj and vr == v and any(p in r for p in prepoznaj):
                kljuc = k
                break
        if kljuc is None:
            kljuc = next((k for k, (naziv, vr, prepoznaj, *_d) in GRAVIDNOST_PROFILI.items()
                          if vr == v and not prepoznaj), None)
        if kljuc is None:
            return None, v, None, None, None
        naziv, _, _, mn, mx, pr = GRAVIDNOST_PROFILI[kljuc]
        pod = self.pod.get("gravidnost2", {}).get(kljuc)
        if pod:
            mn, mx, pr = (int(x) for x in pod)
        return kljuc, naziv, mn, mx, pr

    def gravidnost(self, vrsta, rasa=""):
        """(prosek, koliko dana ranije je dozvoljen upis) – zbog starijih delova programa."""
        _, _, mn, mx, pr = self.profil_gravidnosti(vrsta, rasa)
        return (pr, pr - mn) if pr else (None, 0)

    def vrsta_parenja(self, z):
        majka, _ = self.roditelji_parenja(z)
        g = majka if majka in self.model.A else z["grlo"]
        return self.model.A[g]["vrsta"] if g in self.model.A else ""

    def rasa_parenja(self, z):
        majka, _ = self.roditelji_parenja(z)
        return self.model.A[majka]["rasa"] if majka in self.model.A else ""

    def granice_porodjaja(self, z):
        """(najranije, očekivano, najkasnije) datum porođaja za parenje, ili (None, None, None)."""
        d = parsiraj_datum(z["datum"])
        _, _, mn, mx, pr = self.profil_gravidnosti(self.vrsta_parenja(z), self.rasa_parenja(z))
        if not (d and pr):
            return None, None, None
        return d + timedelta(days=mn), d + timedelta(days=pr), d + timedelta(days=mx)

    def ocekivani_porodjaj(self, z):
        return self.granice_porodjaja(z)[1]

    def upis_moguc_od(self, z):
        """Najraniji normalan dan porođaja – od tada je upis mladih otvoren."""
        return self.granice_porodjaja(z)[0]

    def najkasniji_porodjaj(self, z):
        return self.granice_porodjaja(z)[2]

    def podesavanje_gravidnosti(self, parent=None):
        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title("Gestation length")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, "Gestation length", 16, "bold").grid(row=0, column=0, columnspan=5, sticky="w")
        self.lbl(b, "Entering young opens from the minimum number of days. After the maximum number of days the program reports that the "
                    "due date has passed. The average is used for the expected date.", 9, fg=INK2, wraplength=520,
                 justify="left").grid(row=1, column=0, columnspan=5, sticky="w", pady=(4, 12))
        for c, t in enumerate(("Species / breed", "Minimum", "Maximum", "Average", "")):
            self.lbl(b, t, 9, "bold", fg=INK2).grid(row=2, column=c, sticky="w", padx=(0, 14), pady=(0, 4))
        polja = {}
        for i, (k, (naziv, vr, prep, mn0, mx0, pr0)) in enumerate(GRAVIDNOST_PROFILI.items(), 3):
            _, _, mn, mx, pr = self.profil_gravidnosti(vr, prep[0] if prep else "")
            vs = [tk.StringVar(value=str(x)) for x in (mn, mx, pr)]
            self.lbl(b, naziv, 10, "bold").grid(row=i, column=0, sticky="w", padx=(0, 14), pady=3)
            for c, v in enumerate(vs, 1):
                ttk.Entry(b, textvariable=v, width=7, font=self.f(10)).grid(row=i, column=c, sticky="w", padx=(0, 14))
            self.lbl(b, f"default {mn0}–{mx0}, ~{pr0}", 8, fg="#8A968D").grid(row=i, column=4, sticky="w")
            polja[k] = vs
        self.lbl(b, "Gestation is always determined by the breed of the DAM: Romanov ewe × any ram = 138–143 days; "
                    "other ewe × Romanov ram = as a common sheep.", 8, fg="#8A968D").grid(
            row=19, column=0, columnspan=5, sticky="w", pady=(6, 0))
        err = self.lbl(b, "", 10, "bold", fg=BAD, wraplength=520, justify="left")
        err.grid(row=20, column=0, columnspan=5, sticky="w", pady=(8, 0))

        def sacuvaj(ev=None):
            novo = {}
            for k, vs in polja.items():
                naziv = GRAVIDNOST_PROFILI[k][0]
                try:
                    mn, mx, pr = (int(v.get()) for v in vs)
                except ValueError:
                    err.config(text=f"{naziv}: enter whole numbers of days.")
                    return
                if not (30 <= mn <= pr <= mx <= 500):
                    err.config(text=f"{naziv}: must be minimum ≤ average ≤ maximum (30–500 days).")
                    return
                novo[k] = [mn, mx, pr]
            self.pod["gravidnost2"] = novo
            self.pod.pop("gravidnost", None)
            self.sacuvaj_podesavanja()
            win.grab_release()
            win.destroy()
            self.toast("Gestation length saved.")
            self.osvezi_istoriju()

        def vrati():
            for k, vs in polja.items():
                for v, x in zip(vs, GRAVIDNOST_PROFILI[k][3:]):
                    v.set(str(x))

        fb = tk.Frame(b, bg=SURF)
        fb.grid(row=21, column=0, columnspan=5, sticky="ew", pady=(16, 0))
        self.btn(fb, "Defaults", vrati).pack(side="left")
        self.btn(fb, "Save", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        win.bind("<Return>", sacuvaj)
        win.bind("<Escape>", lambda e: win.destroy())
        self.postavi_prozor(win, parent)

    def ceka_porodjaj(self, z):
        return z["dogadjaj"] == PARENJE and not z.get("ishod") and not z["automatski"]

    def opis_ishoda(self, z):
        if z["dogadjaj"] != PARENJE or z["automatski"]:
            return ""
        ish = z.get("ishod", "")
        if ish.startswith("rodjeno:"):
            return "☑ born: " + ish[8:].replace(",", ", ")
        lg = self.leglo(z)
        if lg:
            uk, dd, vec = lg
            return f"🏷 {uk} born {dd} – waiting for ear tag" + (f" ({len(vec)}/{uk})" if vec else "")
        if ish == "uklonjeno":
            return "✕ young not entered in the list"
        if ish == "neuspelo":
            return "✕ failed"
        if ish == "ponovljeno":
            return "– repeated mating"
        oc = self.ocekivani_porodjaj(z)
        od = self.upis_moguc_od(z)
        if od and date.today() < od:
            return f"⏳ pregnant until {oc.strftime('%d/%m/%Y')}"
        return "☐ " + (f"entry open (due {oc.strftime('%d/%m/%Y')})" if oc else "waiting for young")

    def cekaju_porodjaj(self, x=None):
        """Parenja bez upisanih mladih – za svaku ženku samo poslednje parenje."""
        posl = {}
        for z in self.ist.zapisi:
            if z["dogadjaj"] != PARENJE or z["automatski"]:
                continue
            majka, _ = self.roditelji_parenja(z)
            if x is not None and x not in (z["grlo"], majka, self.model.nadji(z.get("partner"))):
                continue
            kljuc = majka or z["grlo"]
            d = parsiraj_datum(z["datum"]) or date.min
            if kljuc not in posl or d >= (parsiraj_datum(posl[kljuc]["datum"]) or date.min):
                posl[kljuc] = z
        lista = [z for z in posl.values() if self.ceka_porodjaj(z)]
        return sorted(lista, key=lambda z: self.ocekivani_porodjaj(z) or date.max)

    def zatvori_stara_parenja(self, nz):
        """Novo parenje iste ženke: ranija parenja bez mladih označi kao ponovljena."""
        if nz.get("dogadjaj") != PARENJE:
            return
        majka, _ = self.roditelji_parenja(nz)
        d = parsiraj_datum(nz["datum"]) or date.min
        for z in self.ist.zapisi:
            if z is not nz and self.ceka_porodjaj(z) and self.roditelji_parenja(z)[0] == majka \
                    and (parsiraj_datum(z["datum"]) or date.min) <= d:
                z["ishod"] = "ponovljeno"

    def oznaci_neuspelo(self, z):
        if messagebox.askyesno("Mating failed",
                               f"Mark the mating of {z['datum']} ({z['grlo']}) as failed?\n"
                               "It will no longer be shown among expected births."):
            z["ishod"] = "neuspelo"
            self.promena("Mating marked as failed.")
            self.osvezi_istoriju()

    def osvezi_istoriju(self):
        self.osvezi_nav_brojac()
        if getattr(self, "ist_win", None) and self.ist_win.winfo_exists() and getattr(self.ist_win, "grlo", None):
            self.istorija_prozor(self.ist_win.grlo)
        if self.strana == "istorija":
            self.prikazi("istorija")

    # ---------- leglo: rođenje se beleži odmah (broj mladih), a u listu grla se upisuju kasnije (minđuša)
    @staticmethod
    def leglo(z):
        """'leglo:N:dd.mm.gggg:id1,id2' -> (N, datum, [upisani]) ili None."""
        ish = z.get("ishod", "") or ""
        if not ish.startswith("leglo:"):
            return None
        delovi = (ish.split(":", 3) + ["", "", "", ""])[:4]
        try:
            n = int(delovi[1])
        except ValueError:
            n = 0
        return n, delovi[2], [i for i in delovi[3].split(",") if i]

    def cekaju_mindjusu(self, x=None):
        rez = []
        for z in self.ist.zapisi:
            if z["dogadjaj"] != PARENJE or not self.leglo(z):
                continue
            majka, otac = self.roditelji_parenja(z)
            if x is not None and x not in (z["grlo"], majka, otac):
                continue
            rez.append(z)
        return sorted(rez, key=lambda z: parsiraj_datum(self.leglo(z)[1]) or date.min)

    def upis_podmlatka(self, z, parent=None, prevremeni=False):
        """Štikliran 'Dobijeni mladi': ako mladi već čekaju minđušu – upis u listu, inače beleženje rođenja."""
        m = self.model
        if z.get("ishod", "").startswith("rodjeno:"):
            ids = [i for i in z["ishod"][8:].split(",") if i]
            messagebox.showinfo("Young stock", "Young have already been entered for this mating: " + ", ".join(ids),
                                parent=parent or self.root)
            if ids and ids[0] in m.A:
                self.karton(ids[0])
            return
        if self.leglo(z):
            self.upis_u_listu(z, parent)
            return
        od = self.upis_moguc_od(z)
        oc = self.ocekivani_porodjaj(z)
        if od and date.today() < od and not prevremeni:
            dana_ost = (od - date.today()).days
            if not messagebox.askyesno(
                    "Pregnancy still in progress",
                    f"Mating: {z['datum']}\nExpected birth: {oc.strftime('%d/%m/%Y')}\n"
                    f"Young can be entered from {od.strftime('%d/%m/%Y')} (in {dana_ost} days).\n\n"
                    "Was this a PREMATURE birth and do you still want to enter the young?",
                    icon="warning", parent=parent or self.root):
                return
            prevremeni = True
        self.rodjenje_dijalog(z, parent, prevremeni)

    def podaci_legla(self, z):
        m = self.model
        majka, otac = self.roditelji_parenja(z)
        ref = majka if majka in m.A else z["grlo"]
        vrsta = m.A.get(ref, {}).get("vrsta") or "cattle"
        rm = (m.A[majka]["rasa"] if majka in m.A else "") or ""
        ro = (m.A[otac]["rasa"] if otac in m.A else "") or ""
        rasa = rasa_potomka(ro, rm)
        return majka, otac, vrsta, rasa, nazivi(vrsta)

    def zaglavlje_legla(self, b, z, majka, otac):
        m = self.model
        top = tk.Frame(b, bg=SURF)
        top.pack(anchor="w", pady=(12, 2))
        for naslov, g, nep in (("Dam", majka, "unknown"), ("Sire", otac, "unknown")):
            self.lbl(top, naslov, 9, "bold", fg=INK2).pack(side="left")
            if g:
                self.cip(top, g, siv=g not in m.A or m.A[g]["osnivac"]).pack(side="left", padx=(6, 16))
            else:
                self.lbl(top, nep, 10, fg=INK2).pack(side="left", padx=(6, 16))

    def postavi_prozor(self, win, parent):
        win.update_idletasks()
        r = parent or self.root
        win.geometry(f"+{r.winfo_rootx() + max(0, (r.winfo_width() - win.winfo_width()) // 2)}"
                     f"+{r.winfo_rooty() + max(10, (r.winfo_height() - win.winfo_height()) // 3)}")

    def rodjenje_dijalog(self, z, parent=None, prevremeni=False):
        """Korak 1: beleži se samo da su mladi rođeni, datum i broj – bez pola i oznaka."""
        m = self.model
        majka, otac, vrsta, rasa, n = self.podaci_legla(z)
        od, oc = self.upis_moguc_od(z), self.ocekivani_porodjaj(z)
        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title("Birth of young")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, f"Young born ({n[3]})", 16, "bold").pack(anchor="w")
        self.zaglavlje_legla(b, z, majka, otac)
        self.lbl(b, f"Mating {z['datum']}" + (f"   ·   due {oc.strftime('%d/%m/%Y')}" if oc else ""),
                 9, fg=INK2).pack(anchor="w", pady=(4, 0))

        g = tk.Frame(b, bg=SURF2)
        g.pack(fill="x", pady=(14, 4))
        gi = tk.Frame(g, bg=SURF2)
        gi.pack(fill="x", padx=14, pady=10)
        dv = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        bv = tk.StringVar(value="1")
        mv = tk.StringVar(value="0")
        for i, (nasl, var, w) in enumerate((("Date of birth", dv, 14), (f"Number alive ({n[3]})", bv, 6),
                                            ("Stillborn", mv, 6))):
            f = tk.Frame(gi, bg=SURF2)
            f.grid(row=0, column=i, sticky="w", padx=(0, 22))
            self.lbl(f, nasl, 9, "bold", fg=INK2, bg=SURF2).pack(anchor="w")
            (ttk.Entry(f, textvariable=var, width=w, font=self.f(11)) if i == 0 else
             ttk.Spinbox(f, from_=0 if i == 2 else 1, to=30, textvariable=var, width=w,
                         font=self.f(11))).pack(anchor="w", pady=(3, 0))
        prev_v = tk.BooleanVar(value=prevremeni)
        if od:
            ttk.Checkbutton(b, text=f"Premature birth (before {od.strftime('%d/%m/%Y')})", variable=prev_v,
                            style="Card.TCheckbutton").pack(anchor="w", pady=(6, 0))
        info = tk.Frame(b, bg=OK_BG)
        info.pack(fill="x", pady=(12, 0))
        self.lbl(info, "You don't need to know the sex, tag or ear tag now. The young will wait on the list "
                       "\"Waiting for ear tag\", and you move them to the animal list when they get their ear tags.",
                 9, fg=OK, bg=OK_BG, wraplength=480, justify="left").pack(anchor="w", padx=12, pady=8)
        err = self.lbl(b, "", 10, "bold", fg=BAD, wraplength=500, justify="left")
        err.pack(anchor="w", pady=(8, 0))

        def sacuvaj(odmah=False):
            d = parsiraj_datum(dv.get())
            if not d:
                err.config(text="The date of birth is not valid. Enter it as 15/03/2026.")
                return
            if d > date.today():
                err.config(text="The date of birth cannot be in the future.")
                return
            dp = parsiraj_datum(z["datum"])
            if dp and d <= dp:
                err.config(text=f"The date of birth must be after the mating date ({z['datum']}).")
                return
            if od and d < od and not prev_v.get():
                err.config(text=f"Pregnancy lasts until {oc.strftime('%d/%m/%Y')} – the earliest date of birth is "
                                f"{od.strftime('%d/%m/%Y')}. For an earlier birth tick \"Premature birth\".")
                return
            try:
                zivih, mrtvih = int(bv.get()), int(mv.get() or 0)
            except ValueError:
                err.config(text="The number of young must be a whole number.")
                return
            if zivih < 0 or mrtvih < 0 or zivih + mrtvih == 0:
                err.config(text="Enter how many young were born.")
                return
            dstr = d.strftime("%d/%m/%Y")
            opis = f"Born alive: {zivih}" + (f", stillborn: {mrtvih}" if mrtvih else "")
            if od and d < od:
                opis += f" – premature birth ({(oc - d).days} days before due date)"
            if majka:
                self.ist.dodaj(majka, "Birth", opis, datum=dstr, partner=otac or "", vrednost=str(zivih))
            if otac and otac in m.A and not m.A[otac]["osnivac"]:
                self.ist.dodaj(otac, "Offspring", opis, datum=dstr, partner=majka or "", automatski=True)
            z["ishod"] = f"leglo:{zivih}:{dstr}:" if zivih else "rodjeno:"
            win.grab_release()
            win.destroy()
            self.promena(f"Birth recorded: {zivih} {n[3]} ({majka or '?'} × {otac or '?'})")
            self.osvezi_istoriju()
            if odmah and zivih:
                self.upis_u_listu(z, parent)

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(16, 0))
        self.btn(fb, "Record birth", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Enter in list now", lambda: sacuvaj(True)).pack(side="right", padx=8)
        self.btn(fb, "Cancel", win.destroy).pack(side="left")
        win.bind("<Escape>", lambda ev: win.destroy())
        self.postavi_prozor(win, parent)

    def ukloni_iz_legla(self, z, parent=None):
        """Mladi koji neće biti upisani u listu (uginuli, prodati…) – uklanjaju se sa spiska čekanja."""
        lg = self.leglo(z)
        if not lg:
            return
        ukupno, dstr, vec = lg
        ostalo = ukupno - len(vec)
        majka, otac, vrsta, rasa, n = self.podaci_legla(z)
        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title("Removing young from the list")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, "Remove young that will not be entered", 15, "bold").pack(anchor="w")
        self.zaglavlje_legla(b, z, majka, otac)
        self.lbl(b, f"Born {dstr}   ·   still waiting on the list: {ostalo} {n[3]}"
                    + (f" (entered: {', '.join(vec)})" if vec else ""), 9, fg=INK2).pack(anchor="w", pady=(4, 12))
        g = tk.Frame(b, bg=SURF)
        g.pack(fill="x")
        kv = tk.StringVar(value="1")
        rv = tk.StringVar(value="Died")
        dv = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        for i, (nasl, w) in enumerate((("How many to remove", None), ("Reason", None), ("Date", None))):
            f = tk.Frame(g, bg=SURF)
            f.grid(row=0, column=i, sticky="w", padx=(0, 16))
            self.lbl(f, nasl, 9, "bold", fg=INK2).pack(anchor="w")
            if i == 0:
                ttk.Spinbox(f, from_=1, to=ostalo, textvariable=kv, width=6, font=self.f(11)).pack(anchor="w", pady=(3, 0))
            elif i == 1:
                ttk.Combobox(f, textvariable=rv, values=["Died", "Sold", "Slaughtered", "Given away", "Other"],
                             width=14, font=self.f(10)).pack(anchor="w", pady=(3, 0))
            else:
                ttk.Entry(f, textvariable=dv, width=12, font=self.f(10)).pack(anchor="w", pady=(3, 0))
        sv_ = tk.BooleanVar(value=False)
        ttk.Checkbutton(b, text=f"Remove all remaining ({ostalo})", variable=sv_, style="Card.TCheckbutton",
                        command=lambda: kv.set(str(ostalo)) if sv_.get() else None).pack(anchor="w", pady=(10, 0))
        self.lbl(b, "Note", 9, "bold", fg=INK2).pack(anchor="w", pady=(10, 3))
        nv = tk.StringVar()
        ttk.Entry(b, textvariable=nv, width=50, font=self.f(10)).pack(anchor="w", fill="x")
        err = self.lbl(b, "", 10, "bold", fg=BAD, wraplength=460, justify="left")
        err.pack(anchor="w", pady=(8, 0))

        def sacuvaj(ev=None):
            try:
                k = int(kv.get())
            except ValueError:
                err.config(text="The number must be a whole number.")
                return
            if not 1 <= k <= ostalo:
                err.config(text=f"You can remove from 1 to {ostalo}.")
                return
            d = parsiraj_datum(dv.get())
            if not d:
                err.config(text="The date is not valid. Enter it as 15/03/2026.")
                return
            razlog = rv.get().strip() or "Other"
            novo_uk = ukupno - k
            if novo_uk <= len(vec):
                z["ishod"] = ("rodjeno:" + ",".join(vec)) if vec else "uklonjeno"
            else:
                z["ishod"] = f"leglo:{novo_uk}:{dstr}:{','.join(vec)}"
            opis = f"From the litter of {dstr} removed {k} {n[3] if k > 1 else n[2]} – {razlog.lower()}"
            if nv.get().strip():
                opis += f" ({nv.get().strip()})"
            if majka:
                self.ist.dodaj(majka, AUTO, opis, datum=d.strftime("%d/%m/%Y"), automatski=True)
            win.grab_release()
            win.destroy()
            self.promena(f"Removed from the list: {k} {n[3] if k > 1 else n[2]} ({razlog.lower()})")
            self.osvezi_istoriju()

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(16, 0))
        self.btn(fb, "Remove from list", sacuvaj, "Danger").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        win.bind("<Return>", sacuvaj)
        win.bind("<Escape>", lambda ev: win.destroy())
        self.postavi_prozor(win, parent)

    def upis_u_listu(self, z, parent=None):
        """Korak 2: mladi dobili minđušu – upis oznake, markice i pola; datum i stado su automatski."""
        m = self.model
        lg = self.leglo(z)
        if not lg:
            return
        ukupno, dstr, vec = lg
        d = parsiraj_datum(dstr) or date.today()
        majka, otac, vrsta, rasa, n = self.podaci_legla(z)
        ostalo = max(1, ukupno - len(vec))

        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title(f"Entry in animal list – {n[3]}")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, f"Entry in animal list ({n[3]})", 16, "bold").pack(anchor="w")
        self.zaglavlje_legla(b, z, majka, otac)
        star = (date.today() - d).days
        self.lbl(b, f"Litter: {ukupno} {n[3]}" + (f", already entered {len(vec)} ({', '.join(vec)})" if vec else "")
                 + (f"   ·   inbreeding of young {pct(m.kinship(majka, otac))}" if majka in m.A and otac in m.A else ""),
                 9, fg=INK2).pack(anchor="w", pady=(4, 0))

        auto = tk.Frame(b, bg=SURF2)
        auto.pack(fill="x", pady=(14, 4))
        g1 = tk.Frame(auto, bg=SURF2)
        g1.pack(fill="x", padx=14, pady=10)
        sv = tk.StringVar(value=STADO_PODMLADAK)
        for i, (nasl, vr) in enumerate((("Date of birth", f"{dstr}  ({star} days)"),
                                        ("Species / breed", f"{vrsta}{', ' + rasa if rasa else ''}"))):
            f = tk.Frame(g1, bg=SURF2)
            f.grid(row=0, column=i if i == 0 else 2, sticky="w", padx=(0, 22))
            self.lbl(f, nasl, 9, "bold", fg=INK2, bg=SURF2).pack(anchor="w")
            self.lbl(f, vr, 10, "bold", bg=SURF2).pack(anchor="w", pady=(6, 0))
        f = tk.Frame(g1, bg=SURF2)
        f.grid(row=0, column=1, sticky="w", padx=(0, 22))
        self.lbl(f, "Herd", 9, "bold", fg=INK2, bg=SURF2).pack(anchor="w")
        ttk.Combobox(f, textvariable=sv, values=[STADO_PODMLADAK] + [x for x in self.sva_stada()
                                                                    if x != STADO_PODMLADAK],
                     width=16, font=self.f(10)).pack(anchor="w", pady=(3, 0))
        self.lbl(b, "Date of birth and herd are automatic.", 8, fg="#8A968D").pack(anchor="w")

        self.lbl(b, f"Enter each {n[2]}", 11, "bold").pack(anchor="w", pady=(14, 4))
        zag = tk.Frame(b, bg=SURF)
        zag.pack(fill="x")
        for t, w in (("Tag / name *", 17), ("Ear tag number", 19), ("Sex *", 20)):
            self.lbl(zag, t, 9, "bold", fg=INK2, width=w, anchor="w").pack(side="left")
        lista = tk.Frame(b, bg=SURF)
        lista.pack(fill="x")
        redovi = []

        def dodaj_red(fokus=True):
            r = tk.Frame(lista, bg=SURF)
            r.pack(fill="x", pady=3)
            iv, mv, pv = tk.StringVar(), tk.StringVar(), tk.StringVar()
            e = ttk.Entry(r, textvariable=iv, width=15, font=self.f(10))
            e.pack(side="left", padx=(0, 10))
            ttk.Entry(r, textvariable=mv, width=17, font=self.f(10)).pack(side="left", padx=(0, 10))
            seg = tk.Frame(r, bg=SURF2)
            seg.pack(side="left")
            dug = {}

            def postavi(k):
                pv.set(k)
                for kk, lb in dug.items():
                    lb.config(bg=SURF if kk == k else SURF2, fg=INK if kk == k else INK2)
            for k, t in (("z", "Female"), ("m", "Male")):
                lb = tk.Label(seg, text=t, font=self.f(9, "bold"), padx=12, pady=4, bg=SURF2, fg=INK2,
                              cursor="hand2")
                lb.pack(side="left", padx=2, pady=2)
                lb.bind("<Button-1>", lambda ev, k=k: postavi(k))
                dug[k] = lb
            red_ = {"okvir": r, "id": iv, "mk": mv, "pol": pv}
            x_ = self.lbl(r, "✕", 11, "bold", fg=INK2, cursor="hand2")
            x_.pack(side="left", padx=10)
            x_.bind("<Button-1>", lambda ev: ukloni(red_))
            redovi.append(red_)
            if fokus:
                e.focus_set()

        def ukloni(red_):
            if len(redovi) > 1:
                red_["okvir"].destroy()
                redovi.remove(red_)

        for i in range(ostalo):
            dodaj_red(fokus=i == 0)
        lk = self.lbl(b, f"+ Add another {n[2]}", 9, "bold", fg=BRAND, cursor="hand2")
        lk.pack(anchor="w", pady=(6, 0))
        lk.bind("<Button-1>", lambda e: dodaj_red())
        self.lbl(b, "If a young animal has no ear tag yet, remove its row (✕) – you will enter it later.", 8,
                 fg="#8A968D").pack(anchor="w", pady=(2, 0))
        err = self.lbl(b, "", 10, "bold", fg=BAD, wraplength=560, justify="left")
        err.pack(anchor="w", pady=(10, 0))

        def sacuvaj(ev=None):
            stado = sv.get().strip() or STADO_PODMLADAK
            novi, vidjeni = [], set()
            for i, r in enumerate(redovi, 1):
                nid, mk, pol = r["id"].get().strip(), r["mk"].get().strip(), r["pol"].get()
                if not nid:
                    err.config(text=f"Enter a tag or name for {n[2]} no. {i} (or remove the row ✕).")
                    return
                if nid in vidjeni or self.red_za(nid) or nid in m.A:
                    err.config(text=f"Tag {nid} already exists. Choose another.")
                    return
                if mk:
                    poruka = self.markica_zauzeta(mk, None)
                    if poruka or any(norm(mk) == norm(x["markica"]) for x in novi):
                        err.config(text=poruka or f"Ear tag {mk} is entered twice.")
                        return
                if pol not in ("m", "z"):
                    err.config(text=f"Choose the sex for {nid}.")
                    return
                vidjeni.add(nid)
                novi.append(dict(id=nid, markica=mk, vrsta=vrsta, rasa=rasa, pol=pol, datum_rodjenja=dstr,
                                 otac=otac or "", majka=majka or "", status="aktivno", stado=stado,
                                 poreklo="Home bred", poreklo_od="", stare_markice="", uneto=sada()))
            for x in novi:
                self.tab.redovi.append(x)
                self.ist.dodaj(x["id"], AUTO, f"Born {dstr} from {majka or '?'} × {otac or '?'} – entered in the list "
                                              f"(herd {stado})", automatski=True)
            sad = vec + [x["id"] for x in novi]
            opis = ", ".join(f"{x['id']} ({'male' if x['pol'] == 'm' else 'female'})" for x in novi)
            if majka:
                self.ist.dodaj(majka, AUTO, f"Young stock entered in the list: {opis}", automatski=True)
            if len(sad) >= ukupno:
                z["ishod"] = "rodjeno:" + ",".join(sad)
            else:
                ost = ukupno - len(sad)
                cekaj = messagebox.askyesno(
                    "Remaining young", f"Entered {len(sad)} of {ukupno}. {ost} remaining.\n\n"
                                    "Are the others still waiting for ear tags?\n\n"
                                    "Yes – they stay on the list for later\nNo – they are gone (died, sold…)",
                    parent=win)
                if cekaj:
                    z["ishod"] = f"leglo:{ukupno}:{dstr}:{','.join(sad)}"
                else:
                    z["ishod"] = "rodjeno:" + ",".join(sad)
                    if majka:
                        self.ist.dodaj(majka, AUTO, f"From the litter ({dstr}) {ost} not entered – died or sold",
                                       automatski=True)
            win.grab_release()
            win.destroy()
            self.promena(f"Entered in the list: {', '.join(x['id'] for x in novi)} (herd {stado})")
            self.osvezi_istoriju()

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(16, 0))
        self.btn(fb, "Enter in animal list", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        self.btn(fb, "Remove from list…", lambda: (win.grab_release(), win.destroy(),
                                                    self.ukloni_iz_legla(z, parent)), "Danger").pack(side="left")
        win.bind("<Return>", sacuvaj)
        win.bind("<Escape>", lambda ev: win.destroy())
        self.postavi_prozor(win, parent)

    def s_istorija(self):
        body = self.strana_okvir(scroll=False)
        self.zaglavlje(body, "History", "All events in the herd – from newest to oldest.",
                       [("Gestation", self.podesavanje_gravidnosti, "Secondary"),
                        ("+  Add event", lambda: self.dogadjaj_forma(None), "Primary")])
        self.istorija_lista(body, None)

    def istorija_lista(self, body, x, bez_dodaj=False):
        """Lista zapisa sa filterima (za jedno grlo ili celo stado)."""
        tb = tk.Frame(body, bg=BG)
        tb.pack(fill="x", pady=(0, 12))
        key = "ist_" + (x or "_sve")
        st = self.ui.setdefault(key, {"vrsta": "All events", "auto": x is not None, "q": ""})
        qv = tk.StringVar(value=st["q"])
        if x is None:
            e = ttk.Entry(tb, textvariable=qv, width=22, font=self.f(10))
            e.pack(side="left", padx=(0, 10))
            ph = self.lbl(tb, "⌕  Animal or ear tag…", 10, fg="#98A39B", bg=SURF)
            def ph_show(*_):
                ph.place_forget() if qv.get() else ph.place(in_=e, x=10, rely=0.5, anchor="w")
            ph.bind("<Button-1>", lambda ev: e.focus_set())
            ph_show()
            qv.trace_add("write", ph_show)
        vv = tk.StringVar(value=st["vrsta"])
        cb = ttk.Combobox(tb, textvariable=vv, values=["All events"] + DOGADJAJI + ["Offspring"],
                          state="readonly", width=22, font=self.f(10))
        cb.pack(side="left")
        av = tk.BooleanVar(value=st["auto"])
        ttk.Checkbutton(tb, text="Show record changes too", variable=av).pack(side="left", padx=12)
        kartice = tk.Frame(body, bg=BG)
        kartice.pack(fill="x", pady=(0, 12))
        cekaju = self.cekaju_porodjaj(x)
        if cekaju:
            co, cc = self.card(body, 16, 10)
            co.pack(fill="x", pady=(0, 12))
            hr = tk.Frame(cc, bg=SURF)
            hr.pack(fill="x")
            self.lbl(hr, f"Waiting for young ({len(cekaju)})", 11, "bold").pack(side="left")
            lg = self.lbl(hr, "Gestation…", 9, "bold", fg=BRAND, cursor="hand2")
            lg.pack(side="right")
            lg.bind("<Button-1>", lambda e: self.podesavanje_gravidnosti())
            self.lbl(cc, "Tick \"Young born\" after the pregnancy time – enter date and number (sex and tag later).", 9, fg=INK2).pack(anchor="w", pady=(0, 6))
            danas = date.today()
            for z in cekaju[:5]:
                majka, otac = self.roditelji_parenja(z)
                oc = self.ocekivani_porodjaj(z)
                r = tk.Frame(cc, bg=SURF)
                r.pack(fill="x", pady=2)
                od = self.upis_moguc_od(z)
                otvoreno = od is None or danas >= od
                cv = tk.BooleanVar(value=False)
                cb_ = ttk.Checkbutton(r, text="Young born", variable=cv, style="Card.TCheckbutton",
                                      command=lambda z=z, cv=cv: (cv.set(False), self.upis_podmlatka(z)))
                cb_.pack(side="left")
                if not otvoreno:
                    cb_.state(["disabled"])
                self.lbl(r, f"{majka or '?'} × {otac or '?'}", 10, "bold").pack(side="left", padx=(14, 10))
                self.lbl(r, f"mated {z['datum']}", 9, fg=INK2).pack(side="left")
                if oc:
                    dk = self.najkasniji_porodjaj(z)
                    if not otvoreno:
                        txt, boja = (f"pregnant · birth {od.strftime('%d/%m')}–{dk.strftime('%d/%m/%Y')} "
                                     f"(in {(od - danas).days} d.)"), INK2
                    elif danas <= dk:
                        txt, boja = f"birth ~{oc.strftime('%d/%m')} (latest {dk.strftime('%d/%m')}) – entry open", WARN
                    else:
                        txt, boja = f"due date passed {dk.strftime('%d/%m/%Y')} ({(danas - dk).days} d.)", BAD
                    self.lbl(r, txt, 9, "bold", fg=boja).pack(side="left", padx=12)
                if not otvoreno:
                    lp = self.lbl(r, "prevremeni?", 9, fg=BRAND, cursor="hand2")
                    lp.pack(side="right", padx=(10, 0))
                    lp.bind("<Button-1>", lambda e, z=z: self.upis_podmlatka(z))
                lk = self.lbl(r, "failed", 9, fg=INK2, cursor="hand2")
                lk.pack(side="right")
                lk.bind("<Button-1>", lambda e, z=z: self.oznaci_neuspelo(z))
            if len(cekaju) > 5:
                self.lbl(cc, f"… and {len(cekaju) - 5} more (see the \"Young\" column in the table)", 9, fg=INK2).pack(anchor="w")
        legla = self.cekaju_mindjusu(x)
        if legla:
            lo, lc = self.card(body, 16, 10)
            lo.pack(fill="x", pady=(0, 12))
            self.lbl(lc, f"Waiting for ear tag – entry in animal list ({len(legla)})", 11, "bold").pack(anchor="w")
            self.lbl(lc, "Young born without tag and sex. When they get their ear tags, click \"Enter in list\".", 9,
                     fg=INK2).pack(anchor="w", pady=(0, 6))
            for z in legla:
                uk, dd, vec = self.leglo(z)
                majka, otac = self.roditelji_parenja(z)
                dr = parsiraj_datum(dd)
                star = (date.today() - dr).days if dr else None
                n_ = nazivi(self.vrsta_parenja(z))
                r = tk.Frame(lc, bg=SURF)
                r.pack(fill="x", pady=2)
                self.btn(r, "Enter in list", lambda z=z: self.upis_u_listu(z), "Primary").pack(side="right")
                self.btn(r, "Remove…", lambda z=z: self.ukloni_iz_legla(z), "Danger").pack(side="right", padx=8)
                self.lbl(r, f"{majka or '?'} × {otac or '?'}", 10, "bold").pack(side="left")
                self.lbl(r, f"{uk - len(vec)} {n_[3]}" + (f" (of {uk})" if vec else ""), 10, "bold",
                         fg=BRAND).pack(side="left", padx=12)
                self.lbl(r, f"born {dd}" + (f" · {star} d. old" if star is not None else ""), 9,
                         fg=INK2).pack(side="left")
        dugmad = tk.Frame(body, bg=BG)
        dugmad.pack(fill="x", pady=(0, 10))
        holder = tk.Frame(body, bg=BG)
        holder.pack(fill="both", expand=True)
        stanje = {"tv": None}
        st.setdefault("grupa", "")

        def crtaj_kartice():
            for w in kartice.winfo_children():
                w.destroy()
            svi = self.ist.za(x)
            for i, (naziv, vrste) in enumerate([("All", None)] + GRUPE):
                if x is None and vrste:      # za celo stado porođaj se broji jednom (kod majke)
                    vrste = tuple(v for v in vrste if v != "Offspring")
                zz = [z for z in svi if vrste is None or z["dogadjaj"] in vrste]
                if vrste is None:
                    zz = [z for z in zz if av.get() or not z["automatski"]]
                akt = st["grupa"] == ("" if vrste is None else naziv)
                bg = BRAND if akt else SURF
                fg, fg2 = ("#FFFFFF", "#D7E6DC") if akt else (INK, INK2)
                k = tk.Frame(kartice, bg=bg if akt else LINE, cursor="hand2")
                k.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0))
                kartice.columnconfigure(i, weight=1, uniform="k")
                u = tk.Frame(k, bg=bg)
                u.pack(fill="both", expand=True, padx=1, pady=1)
                l1 = tk.Label(u, text=naziv, font=self.f(10, "bold"), bg=bg, fg=fg)
                l1.pack(anchor="w", padx=14, pady=(10, 0))
                l2 = tk.Label(u, text=str(len(zz)), font=self.f(20, "bold"), bg=bg, fg=fg)
                l2.pack(anchor="w", padx=14)
                if vrste is None:
                    pod = "all records"
                elif zz:
                    z0 = zz[0]
                    pod = f"last {z0['datum']}"
                    if naziv == "Matings" and z0["partner"]:
                        pod += f"\nwith {z0['partner']}"
                    elif naziv == "Treatments" and (z0["vrednost"] or z0["dogadjaj"] == "Vaccination"):
                        det = z0["vrednost"] or "vaccination"
                        pod += "\n" + (det if len(det) <= 24 else det[:23] + "…")
                else:
                    pod = "no records"
                l3 = tk.Label(u, text=pod, font=self.f(9), bg=bg, fg=fg2, justify="left")
                l3.pack(anchor="w", padx=14, pady=(0, 10))
                for w in (k, u, l1, l2, l3):
                    w.bind("<Button-1>", lambda e, n=("" if vrste is None else naziv): izaberi_grupu(n))

        def izaberi_grupu(n):
            st["grupa"] = n
            if n:
                vv.set("All events")
            osvezi()

        def osvezi(*_):
            st.update(vrsta=vv.get(), auto=av.get(), q=qv.get())
            for w in holder.winfo_children():
                w.destroy()
            crtaj_kartice()
            grupa = dict(GRUPE).get(st["grupa"])
            if grupa and x is None:
                grupa = tuple(v for v in grupa if v != "Offspring")
            zap = self.ist.za(x, auto=av.get() or bool(grupa))
            if grupa:
                zap = [z for z in zap if z["dogadjaj"] in grupa]
            if vv.get() != "All events":
                zap = [z for z in zap if z["dogadjaj"] == vv.get()]
            q = norm(qv.get())
            if q and x is None:
                ids = {g for g in self.model.A if q in norm(g) or q in norm(self.model.A[g]["markica"])}
                zap = [z for z in zap if z["grlo"] in ids or q in norm(z["grlo"])]
            if not zap:
                self.napomena(holder, "No records for the selected filter." if self.ist.za(x) else
                              "No records yet. Add the first event with the \"Add event\" button.", ok=True)
                stanje["tv"] = None
                return
            wrap, tv = self.tabela_istorije(holder, zap, x is None)
            wrap.pack(fill="both", expand=True)
            stanje["tv"] = tv
            tv.bind("<Double-1>", lambda e: izmeni())
            tv.bind("<Delete>", lambda e: obrisi())

            def klik_mladi(e, tv=tv):
                if tv.identify_region(e.x, e.y) != "cell":
                    return
                kol = tv.identify_column(e.x)
                if tv.heading(kol, "text") != "Young":
                    return
                iid = tv.identify_row(e.y)
                if iid:
                    z = self.ist.zapisi[int(iid)]
                    if z["dogadjaj"] == PARENJE and not z["automatski"]:
                        self.upis_podmlatka(z)
                        return "break"
            tv.bind("<Button-1>", klik_mladi, add="+")

        def izabran():
            tv = stanje["tv"]
            sel = tv.selection() if tv else ()
            if not sel:
                self.toast("First select a record in the table.")
                return None
            return self.ist.zapisi[int(sel[0])]

        def izmeni():
            z = izabran()
            if not z:
                return
            if z["automatski"]:
                self.toast("Automatic records cannot be changed.")
                return
            self.dogadjaj_forma(z["grlo"], zapis=z, posle=osvezi)

        def obrisi():
            z = izabran()
            if not z:
                return
            if z["automatski"]:
                self.toast("Automatic records cannot be deleted.")
                return
            if messagebox.askyesno("Delete record", f"Delete the record \"{z['dogadjaj']}\" of {z['datum']}?"):
                self.ist.zapisi.remove(z)
                self.ist.sacuvaj()
                osvezi()

        if x and not bez_dodaj:
            self.btn(dugmad, "+  Add event", lambda: self.dogadjaj_forma(x, posle=osvezi), "Primary").pack(
                side="left", padx=(0, 8))
        self.btn(dugmad, "Edit record", izmeni).pack(side="left")
        self.btn(dugmad, "Delete record", obrisi, "Danger").pack(side="left", padx=8)
        self.lbl(dugmad, "Double-click edits a record. Grey = automatic changes.", 9,
                 fg=INK2).pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: (st.update(grupa=""), osvezi()))
        av.trace_add("write", osvezi)
        qv.trace_add("write", osvezi)
        osvezi()
        return osvezi

    def azuriraj_grlo(self, red, novi):
        """Primeni izmene podataka grla (bez promene oznake) i zapiši ih u istoriju. Vrati poruku greške ili None."""
        mk = (novi.get("markica") or "").strip()
        if mk and norm(mk) != norm(red.get("markica")):
            poruka = self.markica_zauzeta(mk, red)
            if poruka:
                return poruka
        d = (novi.get("datum_rodjenja") or "").strip()
        if d:
            pd = parsiraj_datum(d)
            if not pd:
                return "The date of birth is not valid. Enter it as 15/03/2022."
            if pd == parsiraj_datum(red.get("datum_rodjenja")):
                novi["datum_rodjenja"] = red.get("datum_rodjenja", "")   # isti datum, drugi zapis – ne menjaj
            else:
                novi["datum_rodjenja"] = pd.strftime("%d/%m/%Y")
        for p in ("otac", "majka"):
            v = (novi.get(p) or "").strip()
            if v and (v == red["id"] or self.model.nadji(v) == red["id"]):
                return "An animal cannot be its own parent."
            novi[p] = self.model.nadji(v) or v if v else ""
        if norm(mk) != norm(red.get("markica")):
            self.primeni_markicu(red, mk, zapamti=True)
        novi.pop("markica", None)
        nazivi_polja = dict(rasa="breed", pol="sex", datum_rodjenja="date of birth", otac="sire", majka="dam",
                            status="status", stado="stado", poreklo="poreklo", poreklo_od="origin from")
        promene = [f"{n}: {red.get(k) or '–'} → {novi.get(k) or '–'}" for k, n in nazivi_polja.items()
                   if k in novi and (red.get(k) or "") != (novi.get(k) or "")]
        if promene:
            self.ist.dodaj(red["id"], AUTO, "Details changed – " + "; ".join(promene), automatski=True)
        red.update(novi)
        return None

    # ================================================================ REZERVNE KOPIJE
    def folder_kopija(self, druga=False):
        if druga:
            d2 = self.pod.get("kopija2", "")
            return os.path.join(d2, "Farmer – backups") if d2 and os.path.isdir(d2) else None
        return os.path.join(os.path.dirname(os.path.abspath(self.tab.putanja)), "backups")

    def fajlovi_podataka(self):
        return [p for p in (self.tab.putanja, self.ist.putanja, self.pods.putanja) if os.path.isfile(p)]

    def napravi_kopiju(self, vrsta="automatic", naziv=None, tiho=True):
        """Kopira tablicu, istoriju i podsetnike. Vraća folder kopije ili None."""
        fajlovi = self.fajlovi_podataka()
        if not fajlovi:
            return None
        naziv = naziv or datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        cilj = None
        for druga in (False, True):
            koren = self.folder_kopija(druga)
            if not koren:
                continue
            try:
                f = os.path.join(koren, vrsta, naziv)
                os.makedirs(f, exist_ok=True)
                for p in fajlovi:
                    shutil.copy2(p, os.path.join(f, os.path.basename(p)))
                if not druga:
                    cilj = f
                self.ocisti_stare_kopije(os.path.join(koren, vrsta))
            except OSError as e:
                if not tiho:
                    messagebox.showerror("Backup", f"Backup was not made:\n{e}")
        if cilj:
            self.pod["poslednja_kopija"] = sada()
            self.sacuvaj_podesavanja()
            self.osvezi_kopija_label()
        return cilj

    @staticmethod
    def ocisti_stare_kopije(folder, zadrzi=30):
        try:
            sve = sorted(d for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d)))
            for d in sve[:-zadrzi]:
                shutil.rmtree(os.path.join(folder, d), ignore_errors=True)
        except OSError:
            pass

    def dnevna_kopija(self):
        """Jednom dnevno, pri pokretanju: rezervne_kopije/automatske/GGGG-MM-DD."""
        dan = date.today().strftime("%Y-%m-%d")
        koren = self.folder_kopija()
        if not os.path.isdir(os.path.join(koren, "automatic", dan)):
            self.napravi_kopiju("automatic", dan)

    def osvezi_kopija_label(self):
        if not getattr(self, "lbl_kopija", None):
            return
        t = self.pod.get("poslednja_kopija", "")
        try:
            dt = datetime.strptime(t, "%Y-%m-%d %H:%M:%S")
            kad = ("today" if dt.date() == date.today() else "yesterday" if dt.date() == date.today() - timedelta(days=1)
                   else dt.strftime("%d/%m/%Y"))
            self.lbl_kopija.config(text=f"Backup: {kad} {dt.strftime('%H:%M')}", fg=INK2)
        except ValueError:
            self.lbl_kopija.config(text="Backup: not made yet", fg=WARN)

    def spisak_kopija(self):
        rez = []
        for druga in (False, True):
            koren = self.folder_kopija(druga)
            if not koren:
                continue
            for vrsta in ("automatic", "manual"):
                f = os.path.join(koren, vrsta)
                if not os.path.isdir(f):
                    continue
                for d in os.listdir(f):
                    put = os.path.join(f, d)
                    tab = os.path.join(put, os.path.basename(self.tab.putanja))
                    if os.path.isfile(tab):
                        try:
                            n = max(0, sum(1 for l in open(tab, encoding="utf-8-sig", errors="ignore") if l.strip()) - 1)
                        except OSError:
                            n = 0
                        rez.append(dict(put=put, vrsta=vrsta, druga=druga, vreme=os.path.getmtime(tab), grla=n))
        return sorted(rez, key=lambda k: -k["vreme"])

    def otvori_folder(self, folder):
        try:
            os.makedirs(folder, exist_ok=True)
            if sys.platform.startswith("win"):
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}" &')
        except OSError:
            messagebox.showinfo("Folder", folder)

    def izaberi_drugu_lokaciju(self):
        d = filedialog.askdirectory(title="Choose a second location for backups (USB, OneDrive, Google Drive…)")
        if not d:
            return
        self.pod["kopija2"] = d
        self.sacuvaj_podesavanja()
        self.napravi_kopiju("manual", tiho=False)
        self.toast("Backups will also be made in the chosen location.")
        if self.strana == "pregled":
            self.prikazi("pregled")

    def vrati_iz_kopije(self):
        kopije = self.spisak_kopija()
        if not kopije:
            messagebox.showinfo("Restore from backup", "There are no backups yet.")
            return
        win = tk.Toplevel(self.root, bg=SURF)
        win.title("Restore from backup")
        win.transient(self.root)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=24, pady=20)
        self.lbl(b, "Restore from backup", 15, "bold").pack(anchor="w")
        self.lbl(b, "Choose a backup. The current data is first saved as a new backup, "
                    "so nothing is lost.", 9, fg=INK2, wraplength=560, justify="left").pack(anchor="w", pady=(2, 10))
        wrap, tv = self.tabela(b, [("kad", "Date and time", 170), ("vrsta", "Type", 140), ("grla", "Animals", 70, "e"),
                                   ("gde", "Location", 160)], visina=min(12, len(kopije)), tree=False)
        wrap.pack(fill="both", expand=True)
        for i, k in enumerate(kopije):
            tv.insert("", "end", iid=str(i), values=(
                datetime.fromtimestamp(k["vreme"]).strftime("%d/%m/%Y %H:%M"),
                {"automatic": "automatic (daily)", "manual": "manual"}[k["vrsta"]], k["grla"],
                "second location" if k["druga"] else "next to the program"), tags=["odd"] if i % 2 else [])
        tv.selection_set("0")

        def vrati():
            sel = tv.selection()
            if not sel:
                return
            k = kopije[int(sel[0])]
            kad = datetime.fromtimestamp(k["vreme"]).strftime("%d/%m/%Y %H:%M")
            if not messagebox.askyesno("Restore from backup", f"Restore data from the backup of {kad} ({k['grla']} animals)?\n\n"
                                       "The current data will first be saved as a new manual backup.", parent=win):
                return
            self.napravi_kopiju("manual", datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + "_before_restore", tiho=False)
            try:
                for ime in os.listdir(k["put"]):
                    cilj = os.path.join(os.path.dirname(os.path.abspath(self.tab.putanja)), ime)
                    shutil.copy2(os.path.join(k["put"], ime), cilj)
                self.tab.ucitaj()
                self.ist = Istorija(self.tab.putanja)
                self.pods = Podsetnici(self.tab.putanja)
            except Exception as e:
                messagebox.showerror("Restore failed", str(e), parent=win)
                return
            self.model = Model(self.tab.redovi)
            win.grab_release()
            win.destroy()
            self.prikazi(self.strana)
            self.toast(f"Data restored from the backup of {kad}.")

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(14, 0))
        self.btn(fb, "Restore selected backup", vrati, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        self.btn(fb, "Open folder", lambda: self.otvori_folder(self.folder_kopija())).pack(side="left")
        win.bind("<Escape>", lambda e: win.destroy())
        self.postavi_prozor(win, None)

    def kartica_kopija(self, body):
        o, c = self.card(body)
        o.pack(fill="x")
        red_ = tk.Frame(c, bg=SURF)
        red_.pack(fill="x")
        lev = tk.Frame(red_, bg=SURF)
        lev.pack(side="left", fill="x", expand=True)
        self.lbl(lev, "Backups", 13, "bold").pack(anchor="w")
        t = self.pod.get("poslednja_kopija", "")
        try:
            kad = datetime.strptime(t, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y at %H:%M")
        except ValueError:
            kad = "not made yet"
        d2 = self.pod.get("kopija2", "")
        self.lbl(lev, f"The program makes a backup every day and keeps the last 30. Last backup: {kad}.", 9,
                 fg=INK2, wraplength=560, justify="left").pack(anchor="w")
        self.lbl(lev, (f"Second location: {d2}" if d2 and os.path.isdir(d2) else
                       "Tip: also choose a second location (USB or OneDrive) – if the computer breaks, the backup remains."),
                 9, fg=OK if d2 and os.path.isdir(d2) else WARN, wraplength=560, justify="left").pack(anchor="w", pady=(2, 0))
        dd = tk.Frame(c, bg=SURF)
        dd.pack(fill="x", pady=(10, 0))
        self.btn(dd, "Make backup now", lambda: (self.napravi_kopiju("manual", tiho=False),
                                                     self.toast("Backup has been made."),
                                                     self.prikazi("pregled")), "Primary").pack(side="left")
        self.btn(dd, "Restore from backup…", self.vrati_iz_kopije).pack(side="left", padx=8)
        self.btn(dd, "Second location…", self.izaberi_drugu_lokaciju).pack(side="left")
        self.btn(dd, "Open folder", lambda: self.otvori_folder(self.folder_kopija())).pack(side="left", padx=8)

    # ================================================================ TER
    def ciklus_tera(self, vrsta):
        return TER_CIKLUS.get(vrsta_norm(vrsta or ""), 21)

    def tera_se(self, z):
        """Posle parenja ženka se ponovo tera – parenje nije uspelo; nudi se novo parenje."""
        majka, otac = self.roditelji_parenja(z)
        z["ter"] = "da"
        z["ishod"] = "neuspelo"
        if majka:
            self.ist.dodaj(majka, TER, f"In heat again – mating of {z['datum']} failed")
        self.promena(f"{majka or '?'}: mating of {z['datum']} marked as failed (in heat again).")
        self.posle_obavestenja()
        if majka and majka in self.model.A:
            self.dogadjaj_forma(majka, preset={"dogadjaj": PARENJE, "partner": otac or ""})

    def ne_tera_se(self, z):
        z["ter"] = "ne"
        self.promena("Recorded: not in heat – probably pregnant.")
        self.posle_obavestenja()

    def posle_novog_zapisa(self, nz):
        """Zajedničko posle svakog ručno upisanog događaja."""
        self.zatvori_stara_parenja(nz)
        if nz.get("dogadjaj") == TER:
            for z in self.cekaju_porodjaj(nz["grlo"]):
                majka, _ = self.roditelji_parenja(z)
                if majka == nz["grlo"] and (parsiraj_datum(z["datum"]) or date.min) < (parsiraj_datum(nz["datum"]) or date.max):
                    z["ishod"] = "neuspelo"
                    z["ter"] = "da"
                    self.ist.dodaj(majka, AUTO, f"Mating of {z['datum']} marked as failed (the female is in heat again)",
                                   datum=nz["datum"], automatski=True)

    def obavestenja_ter(self, x=None):
        danas = date.today()
        rez = []
        # posle parenja: proveriti ter
        for z in self.cekaju_porodjaj():
            if z.get("ter"):
                continue
            majka, otac = self.roditelji_parenja(z)
            if x is not None and x != majka:
                continue
            dp = parsiraj_datum(z["datum"])
            if not dp or majka not in self.model.A:
                continue
            c = self.ciklus_tera(self.model.A[majka]["vrsta"])
            ds = (danas - dp).days
            if c - 3 <= ds <= c + 7:
                kad = dp + timedelta(days=c)
                rez.append(dict(nivo=1, grlo=majka, naslov=f"Check for heat – day {ds} after mating",
                                opis=f"Mated {z['datum']} with {otac or '?'}. If she is not in heat around {kad.strftime('%d/%m')} "
                                     f"(cycle {c} d.), she is probably pregnant.",
                                datum=kad, akcije=[("Not in heat", lambda z=z: self.ne_tera_se(z), "Primary"),
                                                   ("In heat", lambda z=z: self.tera_se(z), "Secondary")]))
        # ženke koje nisu gravidne: sledeći ter iz poslednjeg zabeleženog tera
        poslednji = {}
        for z in self.ist.zapisi:
            if z["dogadjaj"] == TER and not z["automatski"]:
                d = parsiraj_datum(z["datum"])
                if d and (z["grlo"] not in poslednji or d > poslednji[z["grlo"]]):
                    poslednji[z["grlo"]] = d
        gravidne = {self.roditelji_parenja(z)[0] for z in self.cekaju_porodjaj()}
        for g, d in poslednji.items():
            if (x is not None and x != g) or g in gravidne or g not in self.model.A or not self.model.A[g]["aktivno"]:
                continue
            if any(z["dogadjaj"] == PARENJE and self.roditelji_parenja(z)[0] == g and
                   (parsiraj_datum(z["datum"]) or date.min) >= d for z in self.ist.zapisi):
                continue
            c = self.ciklus_tera(self.model.A[g]["vrsta"])
            sl = d + timedelta(days=c)
            while sl < danas - timedelta(days=2):
                sl += timedelta(days=c)
            if (sl - danas).days <= 3:
                rez.append(dict(nivo=2 if sl > danas else 1, grlo=g,
                                naslov="Heat expected " + ("today" if sl == danas else f"around {sl.strftime('%d/%m')}"),
                                opis=f"Last heat {d.strftime('%d/%m/%Y')} · cycle {c} days. Watch whether she comes into heat.",
                                datum=sl, akcije=[("Record mating", lambda g=g: self.dogadjaj_forma(
                                    g, preset={"dogadjaj": PARENJE}), "Primary"),
                                                  ("In heat", lambda g=g: self.dogadjaj_forma(
                                                      g, preset={"dogadjaj": TER}), "Secondary")]))
        return rez

    # ================================================================ KARENCA
    def karenca(self, x):
        """{'meso': (do_datum, zapis), 'mleko': (...)} – samo karence koje još traju."""
        danas = date.today()
        rez = {}
        for z in self.ist.zapisi:
            if z["grlo"] != x:
                continue
            d = parsiraj_datum(z["datum"])
            if not d:
                continue
            for vrsta, kljuc in (("meso", "karenca_meso"), ("mleko", "karenca_mleko")):
                try:
                    n = int(z.get(kljuc) or 0)
                except ValueError:
                    n = 0
                if n > 0:
                    do = d + timedelta(days=n)
                    if do >= danas and (vrsta not in rez or do > rez[vrsta][0]):
                        rez[vrsta] = (do, z)
        return rez

    def opis_karence(self, x, kratko=False):
        k = self.karenca(x)
        if not k:
            return ""
        delovi = []
        for vrsta in ("meso", "mleko"):
            if vrsta in k:
                do = k[vrsta][0]
                ost = (do - date.today()).days
                delovi.append(f"{dict(meso='meat', mleko='milk')[vrsta]} until {do.strftime('%d/%m')}" + ("" if kratko else f" ({ost} d. left)"))
        return ", ".join(delovi)

    def obavestenja_karenca(self, x=None):
        rez = []
        for a in self.model.grla(True):
            if x is not None and a["id"] != x:
                continue
            k = self.karenca(a["id"])
            if not k:
                continue
            z = max(k.values(), key=lambda t: t[0])[1]
            rez.append(dict(nivo=2, grlo=a["id"], naslov="Withdrawal – " + self.opis_karence(a["id"]),
                            opis=f"After: {z['dogadjaj'].lower()} {z['datum']}" + (f" ({z['vrednost']})" if z["vrednost"] else "")
                                 + ". Do not sell for slaughter / do not deliver milk while the withdrawal period lasts.",
                            datum=min(t[0] for t in k.values()), akcije=[]))
        return rez

    def okvir_karence(self, parent, vars_, vrednosti=("", "")):
        """Dva polja za karencu (meso, mleko) – za lečenje i vakcinaciju."""
        f = tk.Frame(parent, bg=SURF)
        for i, (naslov, kljuc, v0) in enumerate((("Meat withdrawal (days)", "karenca_meso", vrednosti[0]),
                                                ("milk (days)", "karenca_mleko", vrednosti[1]))):
            ff = tk.Frame(f, bg=SURF)
            ff.pack(side="left", padx=(0 if i == 0 else 8, 0))
            self.lbl(ff, naslov, 9, "bold", fg=BAD).pack(anchor="w")
            v = tk.StringVar(value=v0)
            ttk.Entry(ff, textvariable=v, width=7, font=self.f(10)).pack(anchor="w", pady=(3, 0))
            vars_[kljuc] = v
        return f

    @staticmethod
    def broj_dana(v):
        t = "".join(ch for ch in str(v or "") if ch.isdigit())
        return t if t and int(t) > 0 else ""

    # ================================================================ PLODNOST
    def plodnost(self, x):
        import re
        por = sorted((z for z in self.ist.zapisi if z["grlo"] == x and z["dogadjaj"] == "Birth"),
                     key=lambda z: parsiraj_datum(z["datum"]) or date.min)
        zivih = mrtvih = 0
        for z in por:
            v = "".join(ch for ch in z.get("vrednost", "") if ch.isdigit())
            zivih += int(v) if v else 1
            mm = re.search(r"stillborn:\s*(\d+)", z.get("opis", ""))
            mrtvih += int(mm.group(1)) if mm else 0
        datumi = [parsiraj_datum(z["datum"]) for z in por if parsiraj_datum(z["datum"])]
        razmaci = [(b - a).days for a, b in zip(datumi, datumi[1:]) if (b - a).days > 20]
        parenja = [z for z in self.ist.zapisi if z["dogadjaj"] == PARENJE and not z["automatski"]
                   and self.roditelji_parenja(z)[0] == x]
        neuspela = sum(1 for z in parenja if z.get("ishod") == "neuspelo")
        a = self.model.A.get(x, {})
        prvi = None
        if datumi and a.get("datum"):
            prvi = (datumi[0].year - a["datum"].year) * 12 + datumi[0].month - a["datum"].month
        potomaka = sum(1 for pr in self.model.P.values() if x in pr)
        return dict(porodjaja=len(por), zivih=zivih, mrtvih=mrtvih,
                    prosek=(zivih / len(por)) if por else 0,
                    razmak=(sum(razmaci) / len(razmaci)) if razmaci else None,
                    parenja=len(parenja), neuspela=neuspela,
                    po_gravidnosti=(len(parenja) / len(por)) if por and len(parenja) >= len(por) else None,
                    prvi_mes=prvi, potomaka=potomaka, poslednji=datumi[-1] if datumi else None)

    def kartica_plodnosti(self, body):
        m = self.model
        redovi = []
        for a in m.grla(True):
            if a["pol"] != "z":
                continue
            p = self.plodnost(a["id"])
            if p["porodjaja"] or p["parenja"]:
                redovi.append((a, p))
        o, c = self.card(body)
        o.pack(fill="x")
        self.lbl(c, "Female fertility", 13, "bold").pack(anchor="w")
        self.lbl(c, "Sorted by young per birth. Helps you choose which females to keep for breeding.", 9,
                 fg=INK2).pack(anchor="w", pady=(0, 10))
        if not redovi:
            self.lbl(c, "No matings or births recorded yet.", 10, fg=INK2).pack(anchor="w")
            return
        redovi.sort(key=lambda t: (-t[1]["prosek"], -t[1]["porodjaja"]))
        wrap, tv = self.tabela(c, [("id", "Female", 95), ("rasa", "Breed", 100), ("por", "Births", 80, "e"),
                                   ("ziv", "Alive", 60, "e"), ("pros", "Average", 70, "e"),
                                   ("mrt", "Stillborn", 65, "e"), ("raz", "Interval (d.)", 90, "e"),
                                   ("par", "Matings/birth", 105, "e"), ("pos", "Last", 95)],
                               visina=min(10, len(redovi)))
        wrap.pack(fill="x")
        for i, (a, p) in enumerate(redovi):
            tv.insert("", "end", iid=a["id"], text="  " + a["id"], image=self.img_tag, values=(
                a["rasa"] or "–", p["porodjaja"], p["zivih"], f"{p['prosek']:.1f}" if p["porodjaja"] else "–",
                p["mrtvih"] or "–", round(p["razmak"]) if p["razmak"] else "–",
                f"{p['po_gravidnosti']:.1f}" if p["po_gravidnosti"] else "–",
                p["poslednji"].strftime("%d/%m/%Y") if p["poslednji"] else "–"), tags=["odd"] if i % 2 else [])
        tv.bind("<Double-1>", lambda e: self.otvori_iz(tv))

    # ---------- obaveštenja i podsetnici
    def obavestenja(self, x=None):
        """Sva obaveštenja: [{nivo, grlo, naslov, opis, kljuc, akcije:[(tekst, fn, stil)]}]. nivo: 0 hitno, 1 uskoro, 2 info, 3 zakazano."""
        m = self.model
        danas = date.today()
        rez = []
        # 1) porođaji
        for z in self.cekaju_porodjaj():
            majka, otac = self.roditelji_parenja(z)
            if x is not None and x not in (majka, otac, z["grlo"]):
                continue
            oc, od = self.ocekivani_porodjaj(z), self.upis_moguc_od(z)
            if not oc:
                continue
            dana = (oc - danas).days
            dk = self.najkasniji_porodjaj(z)
            par = f"{majka or '?'} × {otac or '?'}, mated {z['datum']}"
            akc = [("Young born", lambda z=z: self.upis_podmlatka(z), "Primary"),
                   ("Failed", lambda z=z: self.oznaci_neuspelo(z), "Secondary")]
            if dk and danas > dk:
                rez.append(dict(nivo=0, grlo=majka, naslov=f"Due date passed {(danas - dk).days} d. ago",
                                opis=f"Latest {dk.strftime('%d/%m/%Y')} · {par}. Enter the young or mark the mating as failed.",
                                datum=dk, akcije=akc))
            elif od and danas >= od:
                rez.append(dict(nivo=1, grlo=majka, naslov="Birth expected " +
                                ("today" if dana == 0 else f"in {dana} d." if dana > 0 else "any day now"),
                                opis=f"Around {oc.strftime('%d/%m/%Y')}, latest {dk.strftime('%d/%m/%Y')} · {par}. "
                                     "Entering young is open.",
                                datum=oc, akcije=akc))
            elif dana <= 21:
                rez.append(dict(nivo=2, grlo=majka, naslov=f"Birth in {dana} d.",
                                opis=f"Due {oc.strftime('%d/%m/%Y')} · {par}. Prepare the birthing area.",
                                datum=oc, akcije=[]))
        # 2) mladi čekaju minđušu
        for z in self.cekaju_mindjusu(x):
            uk, dd, vec = self.leglo(z)
            majka, otac = self.roditelji_parenja(z)
            dr = parsiraj_datum(dd)
            star = (danas - dr).days if dr else 0
            n_ = nazivi(self.vrsta_parenja(z))
            rez.append(dict(nivo=1 if star >= 14 else 2, grlo=majka,
                            naslov=f"{uk - len(vec)} {n_[3]} waiting for ear tag ({star} d.)",
                            opis=f"Born {dd} · {majka or '?'} × {otac or '?'}. Enter them in the animal list when they get their ear tags.",
                            datum=dr, akcije=[("Enter in list", lambda z=z: self.upis_u_listu(z), "Primary"),
                                              ("Remove…", lambda z=z: self.ukloni_iz_legla(z), "Secondary")]))
        # 3) podsetnici
        for p in self.pods.aktivni(x):
            d = parsiraj_datum(p["datum"]) or danas
            dana = (d - danas).days
            nivo = 0 if dana < 0 else 1 if dana <= 3 else 3
            kad = (f"{-dana} d. overdue" if dana < 0 else "today" if dana == 0 else
                   "tomorrow" if dana == 1 else f"in {dana} d.")
            pon = int(p.get("ponavljanje") or 0)
            rez.append(dict(nivo=nivo, grlo=p["grlo"] or None, naslov=f"{p['tip'] or 'Reminder'} – {kad}",
                            opis=f"{p['datum']}" + (f" · {p['tekst']}" if p["tekst"] else "") +
                                 (f" · repeats every {pon} d." if pon else ""),
                            datum=d, podsetnik=p,
                            akcije=[("Done", lambda p=p: self.podsetnik_uradjen(p), "Primary"),
                                    ("Postpone 7 d.", lambda p=p: self.podsetnik_odlozi(p, 7), "Secondary"),
                                    ("Edit", lambda p=p: self.podsetnik_forma(zapis=p), "Secondary")]))
        # 4) ter i karenca
        rez += self.obavestenja_ter(x)
        rez += self.obavestenja_karenca(x)
        # 5) greške u podacima
        if x is None and m.warn:
            rez.append(dict(nivo=2, grlo=None, naslov=f"{len(m.warn)} warnings in the data",
                            opis="The program found possible errors in the pedigree (e.g. wrong sex of a parent or a date).",
                            datum=None, akcije=[("Show", lambda: self.prikazi("pregled"), "Secondary")]))
        rez.sort(key=lambda o: (o["nivo"], o["datum"] or date.max))
        return rez

    def broj_obavestenja(self):
        try:
            return sum(1 for o in self.obavestenja() if o["nivo"] <= 1)
        except Exception:
            return 0

    def osvezi_nav_brojac(self):
        if "obavestenja" not in self.nav:
            return
        n = self.broj_obavestenja()
        row, li, lt = self.nav["obavestenja"]
        lt.config(text="Notifications" + (f"  ({n})" if n else ""))
        li.config(fg=BAD if n else (BRAND if self.strana == "obavestenja" else INK2))

    def red_obavestenja(self, parent, o, bg=SURF, kompakt=False):
        boje = {0: (BAD, BAD_BG, "Urgent"), 1: (WARN, WARN_BG, "Soon"), 2: (BRAND, SURF2, "Info"),
                3: (INK2, SURF2, "Scheduled")}
        fg, bgb, txt = boje[o["nivo"]]
        r = tk.Frame(parent, bg=bg)
        r.pack(fill="x", pady=4)
        tk.Frame(r, bg=fg, width=4).pack(side="left", fill="y", padx=(0, 12))
        lev = tk.Frame(r, bg=bg)
        lev.pack(side="left", fill="x", expand=True)
        gornji = tk.Frame(lev, bg=bg)
        gornji.pack(anchor="w")
        if not kompakt:
            tk.Label(gornji, text=f" {txt} ", font=self.f(8, "bold"), bg=bgb, fg=fg).pack(side="left", padx=(0, 8))
        if o["grlo"] and not kompakt:
            self.cip(gornji, o["grlo"], lambda g=o["grlo"]: self.karton(g), bg=bg,
                     siv=o["grlo"] not in self.model.A or self.model.A[o["grlo"]]["osnivac"]).pack(side="left", padx=(0, 8))
        self.lbl(gornji, o["naslov"], 10, "bold", fg=fg if o["nivo"] == 0 else INK, bg=bg).pack(side="left")
        self.lbl(lev, o["opis"], 9, fg=INK2, bg=bg, wraplength=380 if kompakt else 560,
                 justify="left").pack(anchor="w", pady=(2, 0))
        if kompakt:
            if o["akcije"]:
                ak = tk.Frame(lev, bg=bg)
                ak.pack(anchor="w", pady=(3, 0))
                for t, fn, stil in o["akcije"][:2]:
                    lk = self.lbl(ak, t, 9, "bold", fg=BRAND, bg=bg, cursor="hand2")
                    lk.pack(side="left", padx=(0, 14))
                    lk.bind("<Button-1>", lambda e, fn=fn: fn())
            return r
        if o["akcije"]:
            dd = tk.Frame(r, bg=bg)
            dd.pack(side="right", padx=(10, 0))
            for t, fn, stil in reversed(o["akcije"] if not kompakt else o["akcije"][:1]):
                self.btn(dd, t, fn, stil).pack(side="right", padx=(6, 0))
        return r

    def s_obavestenja(self):
        body = self.strana_okvir()
        sva = self.obavestenja()
        aktivna = [o for o in sva if o["nivo"] <= 2]
        zakazano = [o for o in sva if o["nivo"] == 3]
        n = sum(1 for o in sva if o["nivo"] <= 1)
        self.zaglavlje(body, "Notifications",
                       (f"{n} {'thing needs' if n == 1 else 'things need'} attention." if n else
                        "Nothing urgent.") + " The program tracks births, heat, withdrawal periods, untagged young and reminders.",
                       [("+  New reminder", lambda: self.podsetnik_forma(), "Primary")])
        if not aktivna and not zakazano:
            self.napomena(body, "All is well – no notifications. Add a reminder for an animal (vaccination, "
                                "treatment, sale…) with the \"New reminder\" button.", ok=True)
            return
        for nivo, naslov in ((0, "Urgent"), (1, "Soon"), (2, "For information")):
            grupa = [o for o in aktivna if o["nivo"] == nivo]
            if not grupa:
                continue
            o_, c_ = self.card(body, 18, 12)
            o_.pack(fill="x", pady=(0, 14))
            self.lbl(c_, f"{naslov} ({len(grupa)})", 12, "bold", fg=BAD if nivo == 0 else INK).pack(anchor="w", pady=(0, 4))
            for o in grupa:
                self.red_obavestenja(c_, o)
        if zakazano:
            o_, c_ = self.card(body, 18, 12)
            o_.pack(fill="x", pady=(0, 14))
            self.lbl(c_, f"Scheduled reminders ({len(zakazano)})", 12, "bold").pack(anchor="w", pady=(0, 4))
            for o in zakazano:
                self.red_obavestenja(c_, o)

    def obavestenja_na_startu(self):
        sva = [o for o in self.obavestenja() if o["nivo"] <= 1]
        if not sva:
            return
        win = tk.Toplevel(self.root, bg=SURF)
        win.title("Notifications")
        win.transient(self.root)
        win.resizable(False, False)
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=24, pady=20)
        h = sum(1 for o in sva if o["nivo"] == 0)
        self.lbl(b, f"You have {len(sva)} notifications", 15, "bold").pack(anchor="w")
        self.lbl(b, (f"{h} urgent · " if h else "") + f"{len(sva) - h} soon", 9, fg=BAD if h else INK2).pack(anchor="w", pady=(2, 10))
        for o in sva[:5]:
            r = tk.Frame(b, bg=SURF)
            r.pack(fill="x", pady=3)
            tk.Frame(r, bg=BAD if o["nivo"] == 0 else WARN, width=4, height=34).pack(side="left", padx=(0, 10))
            if o["grlo"]:
                self.cip(r, o["grlo"]).pack(side="left", padx=(0, 8))
            self.lbl(r, o["naslov"], 10, "bold").pack(side="left")
        if len(sva) > 5:
            self.lbl(b, f"… and {len(sva) - 5} more", 9, fg=INK2).pack(anchor="w", pady=(4, 0))
        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(16, 0))
        self.btn(fb, "Open notifications", lambda: (win.destroy(), self.prikazi("obavestenja")), "Primary").pack(side="right")
        self.btn(fb, "Later", win.destroy).pack(side="right", padx=8)
        win.bind("<Escape>", lambda e: win.destroy())
        self.postavi_prozor(win, None)

    def podsetnik_uradjen(self, p):
        d = date.today().strftime("%d/%m/%Y")
        p["gotovo"] = d
        if p["grlo"]:
            tip = p["tip"] if p["tip"] in DOGADJAJI else "Note"
            self.ist.dodaj(p["grlo"], tip, (p["tip"] + ": " if tip == "Note" and p["tip"] not in ("", "Note")
                                            else "") + (p["tekst"] or "done as reminded"), datum=d)
        pon = int(p.get("ponavljanje") or 0)
        if pon:
            sl = (parsiraj_datum(p["datum"]) or date.today()) + timedelta(days=pon)
            while sl <= date.today():
                sl += timedelta(days=pon)
            self.pods.dodaj(p["grlo"], sl.strftime("%d/%m/%Y"), p["tip"], p["tekst"], pon)
        self.promena(f"Done: {p['tip'] or 'reminder'}" + (f" – {p['grlo']}" if p["grlo"] else "") +
                     (" · the next one is scheduled" if pon else ""))
        self.posle_obavestenja()

    def podsetnik_odlozi(self, p, dana):
        d = max(parsiraj_datum(p["datum"]) or date.today(), date.today()) + timedelta(days=dana)
        p["datum"] = d.strftime("%d/%m/%Y")
        self.promena(f"Reminder postponed to {p['datum']}")
        self.posle_obavestenja()

    def posle_obavestenja(self):
        self.osvezi_nav_brojac()
        if getattr(self, "ist_win", None) and self.ist_win.winfo_exists() and getattr(self.ist_win, "grlo", None):
            self.istorija_prozor(self.ist_win.grlo)
        if self.karton_win and self.karton_win.winfo_exists() and getattr(self.karton_win, "grlo", None):
            self.karton(self.karton_win.grlo)

    def podsetnik_forma(self, grlo=None, zapis=None, parent=None):
        m = self.model
        z = zapis or {}
        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title("Edit reminder" if zapis else "New reminder")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, "Edit reminder" if zapis else "New reminder", 16, "bold").pack(anchor="w")
        self.lbl(b, "The program will notify you when the day comes (and 3 days before).", 9, fg=INK2).pack(anchor="w", pady=(2, 10))
        g = tk.Frame(b, bg=SURF)
        g.pack(fill="x")
        gv = tk.StringVar(value=z.get("grlo") or grlo or "")
        dv = tk.StringVar(value=z.get("datum") or (date.today() + timedelta(days=7)).strftime("%d/%m/%Y"))
        f = tk.Frame(g, bg=SURF)
        f.grid(row=0, column=0, sticky="w", padx=(0, 16))
        self.lbl(f, "Animal (empty = whole herd)", 9, "bold", fg=INK2).pack(anchor="w")
        ttk.Combobox(f, textvariable=gv, values=[a["id"] for a in m.grla()], width=18, font=self.f(10)).pack(anchor="w", pady=(3, 0))
        f = tk.Frame(g, bg=SURF)
        f.grid(row=0, column=1, sticky="w")
        self.lbl(f, "Date", 9, "bold", fg=INK2).pack(anchor="w")
        de = tk.Frame(f, bg=SURF)
        de.pack(anchor="w", pady=(3, 0))
        ttk.Entry(de, textvariable=dv, width=12, font=self.f(10)).pack(side="left")
        for t, dn in (("+7", 7), ("+30", 30), ("+90", 90)):
            lk = self.lbl(de, t, 9, "bold", fg=BRAND, cursor="hand2")
            lk.pack(side="left", padx=(8, 0))
            lk.bind("<Button-1>", lambda e, dn=dn: dv.set((date.today() + timedelta(days=dn)).strftime("%d/%m/%Y")))
        self.lbl(b, "What needs to be done", 9, "bold", fg=INK2).pack(anchor="w", pady=(12, 4))
        tv = tk.StringVar(value=z.get("tip", "Vaccination"))
        self.izbor_dogadjaja(b, tv, po_redu=5).pack(anchor="w")
        self.lbl(b, "Description", 9, "bold", fg=INK2).pack(anchor="w", pady=(10, 3))
        xv = tk.StringVar(value=z.get("tekst", ""))
        ttk.Entry(b, textvariable=xv, width=60, font=self.f(10)).pack(anchor="w", fill="x")
        self.lbl(b, "e.g. foot-and-mouth vaccine, check after treatment, sale to a buyer…", 8, fg="#8A968D").pack(anchor="w")
        pf = tk.Frame(b, bg=SURF)
        pf.pack(anchor="w", pady=(10, 0))
        self.lbl(pf, "Repeat", 9, "bold", fg=INK2).pack(side="left", padx=(0, 8))
        opc = {"does not repeat": 0, "every week": 7, "every month": 30, "every 3 months": 90,
               "every 6 months": 182, "every year": 365}
        pon0 = int(z.get("ponavljanje") or 0)
        pv = tk.StringVar(value=next((k for k, v in opc.items() if v == pon0), f"every {pon0} days" if pon0 else "does not repeat"))
        ttk.Combobox(pf, textvariable=pv, values=list(opc), width=16, font=self.f(10)).pack(side="left")
        err = self.lbl(b, "", 10, "bold", fg=BAD)
        err.pack(anchor="w", pady=(8, 0))

        def sacuvaj(ev=None):
            gr = gv.get().strip()
            if gr:
                g2 = m.nadji(gr)
                if not g2:
                    err.config(text=f"Animal \"{gr}\" was not found.")
                    return
                gr = g2
            d = parsiraj_datum(dv.get())
            if not d:
                err.config(text="The date is not valid. Enter it as 15/03/2026.")
                return
            tip = tv.get().strip()
            if not tip:
                err.config(text="Choose what needs to be done.")
                return
            pt = pv.get().strip()
            if pt in opc:
                pon = opc[pt]
            else:
                cif = "".join(ch for ch in pt if ch.isdigit())
                pon = int(cif) if cif else 0
            if zapis:
                zapis.update(grlo=gr, datum=d.strftime("%d/%m/%Y"), tip=tip, tekst=xv.get().strip(), ponavljanje=str(pon))
            else:
                self.pods.dodaj(gr, d.strftime("%d/%m/%Y"), tip, xv.get().strip(), pon)
            win.grab_release()
            win.destroy()
            self.promena(f"Reminder saved for {d.strftime('%d/%m/%Y')}")
            self.posle_obavestenja()

        fb = tk.Frame(b, bg=SURF)
        fb.pack(fill="x", pady=(16, 0))
        if zapis:
            def obrisi():
                if messagebox.askyesno("Delete reminder", "Delete this reminder?", parent=win):
                    self.pods.zapisi.remove(zapis)
                    win.grab_release()
                    win.destroy()
                    self.promena("Reminder deleted.")
                    self.posle_obavestenja()
            self.btn(fb, "Delete", obrisi, "Danger").pack(side="left")
        self.btn(fb, "Save", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        win.bind("<Return>", sacuvaj)
        win.bind("<Escape>", lambda e: win.destroy())
        self.postavi_prozor(win, parent)

    def izbor_dogadjaja(self, parent, var, bg=SURF, po_redu=5):
        """Dugmad za izbor vrste događaja (+ 'Drugo…' za sopstveni naziv). Vraća okvir."""
        okvir = tk.Frame(parent, bg=bg)
        cipovi = {}
        drugo_v = tk.StringVar(value=var.get() if var.get() and var.get() not in DOGADJAJI else "")
        drugo_e = ttk.Entry(okvir, textvariable=drugo_v, width=26, font=self.f(10))
        stanje = {"drugo": bool(drugo_v.get())}

        def stil(*_):
            izbor = var.get()
            for naziv, lb in cipovi.items():
                akt = (naziv == izbor) if naziv != "Other…" else stanje["drugo"]
                lb.config(bg=BRAND if akt else SURF2, fg="#FFFFFF" if akt else INK)
            if stanje["drugo"]:
                drugo_e.grid(row=99, column=0, columnspan=po_redu, sticky="w", pady=(6, 0))
            else:
                drugo_e.grid_remove()

        def klik(naziv):
            if naziv == "Other…":
                stanje["drugo"] = True
                var.set(drugo_v.get().strip())
                stil()
                drugo_e.focus_set()
            else:
                stanje["drugo"] = False
                var.set(naziv)
                stil()

        for i, naziv in enumerate(DOGADJAJI + ["Other…"]):
            lb = tk.Label(okvir, text=naziv, font=self.f(9, "bold"), padx=10, pady=5, cursor="hand2", bg=SURF2, fg=INK)
            lb.grid(row=i // po_redu, column=i % po_redu, sticky="ew", padx=(0, 6), pady=(0, 6))
            lb.bind("<Button-1>", lambda e, n=naziv: klik(n))
            cipovi[naziv] = lb
        drugo_v.trace_add("write", lambda *_: stanje["drugo"] and var.set(drugo_v.get().strip()))
        var.trace_add("write", stil)
        ph = self.lbl(okvir, "e.g. Castration, Weaning…", 9, fg="#98A39B", bg=SURF)
        def ph_show(*_):
            if stanje["drugo"] and not drugo_v.get():
                ph.place(in_=drugo_e, x=8, rely=0.5, anchor="w")
            else:
                ph.place_forget()
        drugo_v.trace_add("write", ph_show)
        var.trace_add("write", ph_show)
        ph.bind("<Button-1>", lambda e: drugo_e.focus_set())
        stil()
        ph_show()
        return okvir

    def istorija_prozor(self, x):
        """Prozor grla: levo podaci (čitanje i izmena), desno unos novog zapisa i cela istorija."""
        m = self.model
        if x not in m.A or m.A[x]["osnivac"]:
            if x in m.A:
                self.karton(x)
            return
        if getattr(self, "ist_win", None) and self.ist_win.winfo_exists():
            win = self.ist_win
            for w in win.winfo_children():
                w.destroy()
        else:
            win = tk.Toplevel(self.root, bg=BG)
            win.geometry(f"1200x780+{self.root.winfo_rootx() + 20}+{max(0, self.root.winfo_rooty() - 10)}")
            win.minsize(1000, 640)
            win.transient(self.root)
            win.bind("<Escape>", lambda e: win.destroy())
            self.ist_win = win
        win.title(f"Animal {x} – details and history")
        win.grlo = x
        a = m.A[x]
        red = self.red_za(x)

        # ---------- zaglavlje
        h = tk.Frame(win, bg=BG)
        h.pack(fill="x", padx=24, pady=(18, 12))
        self.cip(h, x, lambda: self.karton(x), big=True, bg=BG).pack(side="left")
        t = tk.Frame(h, bg=BG)
        t.pack(side="left", padx=14)
        self.lbl(t, f"{m.pol_naziv(x).capitalize()}{', ' + a['rasa'] if a['rasa'] else ''}", 15, "bold").pack(anchor="w")
        st = starost_meseci(a["datum"])
        self.lbl(t, "   ·   ".join(v for v in (
            f"ear tag {a['markica'] or '–'}",
            f"{fmt_starost(st)}" if st is not None else "",
            "in herd" if a["aktivno"] else "deleted" if a.get("obrisano") else "out of herd",
            f"inbreeding {pct(m.inbr(x))}") if v), 10, fg=INK2).pack(anchor="w")
        dd = tk.Frame(h, bg=BG)
        dd.pack(side="right")
        self.btn(dd, "Card", lambda: self.karton(x)).pack(side="left", padx=(0, 8))
        self.btn(dd, "Pedigree – tree", lambda: self.stablo(x)).pack(side="left", padx=(0, 8))
        self.btn(dd, "Partner search", lambda: (win.destroy(), self.idi_na_pretragu(x))).pack(side="left")

        telo = tk.Frame(win, bg=BG)
        telo.pack(fill="both", expand=True, padx=24, pady=(0, 20))

        # ---------- levo: podaci o grlu
        lo, lc = self.card(telo, 18, 16)
        lo.pack(side="left", fill="y", padx=(0, 16))
        self.lbl(lc, "Animal details", 13, "bold").grid(row=0, column=0, columnspan=2, sticky="w")
        self.lbl(lc, "Edit a field and click \"Save details\".", 9, fg=INK2).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))
        sva = m.grla()
        pv = {}

        def polje(r, c, naslov, key, vrednost, vals=None, state="normal", span=1):
            f = tk.Frame(lc, bg=SURF)
            f.grid(row=r, column=c, columnspan=span, sticky="new", padx=(0 if c == 0 else 10, 0), pady=4)
            self.lbl(f, naslov, 9, "bold", fg=INK2).pack(anchor="w")
            v = tk.StringVar(value=vrednost or "")
            w = (ttk.Combobox(f, textvariable=v, values=vals, state=state, width=16, font=self.f(10))
                 if vals is not None else ttk.Entry(f, textvariable=v, width=18, font=self.f(10)))
            w.pack(fill="x", pady=(3, 0))
            pv[key] = v
            return w

        n = nazivi(a["vrsta"])
        pol_opcije = {n[1].capitalize(): "z", n[0].capitalize(): "m"}
        dat = parsiraj_datum(red.get("datum_rodjenja"))
        polje(2, 0, "Ear tag", "markica", red.get("markica"))
        polje(2, 1, "Breed", "rasa", red.get("rasa"), sorted({g["rasa"] for g in sva if g["rasa"]}))
        polje(3, 0, "Sex", "pol", next((k for k, v in pol_opcije.items() if v == red.get("pol")), ""),
              list(pol_opcije), "readonly")
        polje(3, 1, "Date of birth", "datum_rodjenja", dat.strftime("%d/%m/%Y") if dat else red.get("datum_rodjenja"))
        polje(4, 0, "Sire", "otac", red.get("otac"),
              sorted((g["id"] for g in m.A.values() if g["pol"] == "m" and g["id"] != x), key=sort_key))
        polje(4, 1, "Dam", "majka", red.get("majka"),
              sorted((g["id"] for g in m.A.values() if g["pol"] == "z" and g["id"] != x), key=sort_key))
        polje(5, 0, "Herd", "stado", red.get("stado"), self.sva_stada())
        polje(5, 1, "Origin", "poreklo", red.get("poreklo"), POREKLA)
        polje(6, 0, "From whom / where", "poreklo_od", red.get("poreklo_od"),
              sorted({g["poreklo_od"] for g in sva if g["poreklo_od"]}), span=2)
        akt = tk.BooleanVar(value=red.get("status", "aktivno") == "aktivno")
        if red.get("status") != "obrisano":
            ttk.Checkbutton(lc, text="In herd (not sold or dead)", variable=akt,
                            style="Card.TCheckbutton").grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))
        # informacije samo za čitanje
        info = tk.Frame(lc, bg=SURF2)
        info.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        o, mj = m.P[x]
        pot = [k for k, pr in m.P.items() if x in pr]
        for i, (k, v) in enumerate((("Parents", f"{o or '?'} × {mj or '?'}"),
                                    ("Offspring", str(len(pot))),
                                    ("Known ancestors", str(len(m.putanje(x)))),
                                    ("Inbreeding", pct(m.inbr(x))),
                                    ("Added", fmt_upis(a["uneto"])),
                                    ("Previous ear tags", ", ".join(a["stare"]) or "–"))):
            self.lbl(info, k, 9, fg=INK2, bg=SURF2).grid(row=i, column=0, sticky="w", padx=(12, 8), pady=2)
            self.lbl(info, v, 9, "bold", bg=SURF2, wraplength=180, justify="left").grid(
                row=i, column=1, sticky="w", padx=(0, 12), pady=2)
        err_p = self.lbl(lc, "", 9, "bold", fg=BAD, wraplength=330, justify="left")
        err_p.grid(row=9, column=0, columnspan=2, sticky="w", pady=(8, 0))

        def sacuvaj_podatke():
            novi = {k: v.get().strip() for k, v in pv.items()}
            novi["pol"] = pol_opcije.get(novi["pol"], red.get("pol", ""))
            novi["rasa"] = norm_rasa(novi["rasa"])
            o_, m_ = self.model.nadji(novi.get("otac")), self.model.nadji(novi.get("majka"))
            if o_ and m_ and (not novi["rasa"] or novi["rasa"].startswith("hybrid") or novi["rasa"] == "crossbred"):
                novi["rasa"] = rasa_potomka(self.model.A[o_]["rasa"], self.model.A[m_]["rasa"]) or novi["rasa"]
            if red.get("status") != "obrisano":
                novi["status"] = "aktivno" if akt.get() else "neaktivno"
            greska = self.azuriraj_grlo(red, novi)
            if greska:
                err_p.config(text=greska)
                return
            self.promena(f"Details saved: {x}")
            self.istorija_prozor(x)

        self.btn(lc, "Save details", sacuvaj_podatke, "Primary").grid(
            row=10, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        # ---------- desno: novi zapis + istorija
        dsf = ScrollFrame(telo)
        dsf.pack(side="left", fill="both", expand=True)
        desno = tk.Frame(dsf.inner, bg=BG)
        desno.pack(fill="both", expand=True, padx=(0, 10))
        no, nc = self.card(desno, 18, 14)
        no.pack(fill="x", pady=(0, 14))
        self.lbl(nc, "New history record", 13, "bold").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        ev = {}

        def ep(r, c, naslov, key, vrednost="", vals=None, width=16, span=1, state="normal"):
            f = tk.Frame(nc, bg=SURF)
            f.grid(row=r, column=c, columnspan=span, sticky="new", padx=(0 if c == 0 else 10, 0), pady=3)
            lab = self.lbl(f, naslov, 9, "bold", fg=INK2)
            lab.pack(anchor="w")
            v = tk.StringVar(value=vrednost)
            w = (ttk.Combobox(f, textvariable=v, values=vals, width=width, font=self.f(10), state=state)
                 if vals is not None else ttk.Entry(f, textvariable=v, width=width, font=self.f(10)))
            w.pack(fill="x", pady=(3, 0))
            ev[key] = v
            return lab

        for c in range(4):
            nc.columnconfigure(c, weight=1)
        # vrsta događaja – dugmad; podrazumevano prema izabranoj kartici
        grupa = self.ui.get("ist_" + x, {}).get("grupa", "")
        pocetni = {"Matings": "Mating / insemination", "Treatments": "Treatment", "Births": "Birth"}.get(grupa, "")
        ev["dogadjaj"] = tk.StringVar(value=pocetni)
        self.lbl(nc, "Event – click a type", 9, "bold", fg=INK2).grid(row=1, column=0, columnspan=4, sticky="w")
        self.izbor_dogadjaja(nc, ev["dogadjaj"]).grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 2))
        ep(3, 0, "Date", "datum", date.today().strftime("%d/%m/%Y"), width=12)
        suprotni = [k["id"] for k in m.grla() if k["vrsta"] == a["vrsta"] and k["pol"] and k["pol"] != a["pol"]]
        lab_par = ep(3, 1, "Partner", "partner", "", suprotni, width=12)
        f_par_ = lab_par.master
        f_kar = self.okvir_karence(nc, ev)
        lab_vr = ep(3, 2, "Value", "vrednost", "", width=16, span=2)
        ep(4, 0, "Description / note", "opis", "", width=60, span=4)
        poruka = self.lbl(nc, "", 9, fg=INK2, wraplength=620, justify="left")
        poruka.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))
        nazivi_vr = {"Weighing": "Weight (kg)", "Vaccination": "Vaccine / product",
                     "Treatment": "Drug / diagnosis", "Sale": "Buyer / price", "Move": "New herd",
                     "Birth": "Number of young", "Mating / insemination": "Natural / AI"}

        def prilagodi(*_):
            d = ev["dogadjaj"].get()
            lab_vr.config(text=nazivi_vr.get(d, "Value"))
            if d in ("Treatment", "Vaccination"):
                f_par_.grid_remove()
                f_kar.grid(row=3, column=1, sticky="nw", padx=(10, 0), pady=3)
            else:
                f_kar.grid_remove()
                f_par_.grid()
            lab_par.config(text="Partner (bull / ram…)" if d in ("Mating / insemination", "Birth") else "Partner")
            p_ = m.nadji(ev["partner"].get())
            if not d:
                poruka.config(text="Choose the event type by clicking one of the buttons.", fg=INK2)
            elif d == "Mating / insemination" and p_ and p_ != x:
                f = m.kinship(x, p_)
                t_, boja = procena(f, self.prag)
                dp = parsiraj_datum(ev["datum"].get())
                zenka = x if a["pol"] == "z" else p_
                _, nz_, mn_, mx_, pr_ = self.profil_gravidnosti(a["vrsta"], m.A[zenka]["rasa"] if zenka in m.A else "")
                oc = (f"  Expected birth: {(dp + timedelta(days=pr_)).strftime('%d/%m/%Y')} "
                      f"(between {(dp + timedelta(days=mn_)).strftime('%d/%m')} and "
                      f"{(dp + timedelta(days=mx_)).strftime('%d/%m/%Y')}) – gestation: {nz_.lower()}, "
                      f"by the dam's breed." if dp and pr_ else "")
                poruka.config(text=f"Expected inbreeding of offspring: {pct(f)} ({t_}), shared genes "
                                   f"{pct(m.r(x, p_))}.{oc}", fg=boja)
            elif d == "Sale" and "meso" in self.karenca(x):
                poruka.config(text=f"WARNING: meat withdrawal lasts until {self.karenca(x)['meso'][0].strftime('%d/%m/%Y')}! "
                                   "The animal will be marked as out of herd.", fg=BAD)
            elif d in ("Sale", "Death") and a["aktivno"]:
                poruka.config(text="The animal will be marked as out of herd.", fg=WARN)
            elif d in ("Treatment", "Vaccination"):
                poruka.config(text="Enter the withdrawal period from the drug leaflet – the program will warn you while it lasts.", fg=INK2)
            elif d == TER:
                poruka.config(text="If the female has been mated, that mating will be marked as failed.", fg=INK2)
            elif d == "Move":
                poruka.config(text="Enter the new herd in the \"New herd\" field – the animal's herd will change.", fg=INK2)
            else:
                poruka.config(text="")
        for k in ("dogadjaj", "partner", "datum"):
            ev[k].trace_add("write", prilagodi)
        prilagodi()

        def dodaj_zapis(e=None):
            d = parsiraj_datum(ev["datum"].get())
            if not d:
                poruka.config(text="The date is not valid. Enter it as 15/03/2026.", fg=BAD)
                return
            dog = ev["dogadjaj"].get().strip()
            if not dog:
                poruka.config(text="Choose the event type (mating, treatment, birth…).", fg=BAD)
                return
            par = ev["partner"].get().strip()
            par = m.nadji(par) or par
            vr = ev["vrednost"].get().strip()
            leci = dog in ("Treatment", "Vaccination")
            nz = self.ist.dodaj(x, dog, opis=ev["opis"].get().strip(), datum=d.strftime("%d/%m/%Y"),
                                partner="" if leci else par, vrednost=vr,
                                karenca_meso=self.broj_dana(ev["karenca_meso"].get()) if leci else "",
                                karenca_mleko=self.broj_dana(ev["karenca_mleko"].get()) if leci else "")
            self.posle_novog_zapisa(nz)
            if dog in ("Sale", "Death") and red.get("status") == "aktivno":
                red["status"] = "neaktivno"
                self.ist.dodaj(x, AUTO, f"Status: out of herd ({dog.lower()})", automatski=True)
            if dog == "Move" and vr and vr != red.get("stado"):
                staro = red.get("stado") or "–"
                vr = next((s_ for s_ in self.sva_stada() if s_.lower() == vr.lower()), vr)
                red["stado"] = vr
                self.ist.dodaj(x, AUTO, f"Herd: {staro} → {vr}", automatski=True)
            self.promena(f"Recorded: {x} – {dog}")
            self.istorija_prozor(x)

        self.btn(nc, "+  Add record", dodaj_zapis, "Primary").grid(row=5, column=3, sticky="e", pady=(8, 0))

        lista = tk.Frame(desno, bg=BG)
        lista.pack(fill="both", expand=True)
        self.lbl(lista, "History", 13, "bold").pack(anchor="w", pady=(0, 8))
        win.osvezi = self.istorija_lista(lista, x, bez_dodaj=True)
        win.lift()

    def dogadjaj_forma(self, x=None, zapis=None, parent=None, posle=None, preset=None):
        m = self.model
        win = tk.Toplevel(parent or self.root, bg=SURF)
        win.title("Edit event" if zapis else "New event")
        win.transient(parent or self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, "Edit event" if zapis else "New event", 16, "bold").grid(row=0, column=0, columnspan=2,
                                                                                 sticky="w", pady=(0, 12))
        sva = m.grla()
        vars_ = {}

        def polje(r, c, naslov, key, vrednost="", vals=None, state="normal", width=26, span=1):
            f = tk.Frame(b, bg=SURF)
            f.grid(row=r, column=c, columnspan=span, sticky="new", padx=(0 if c == 0 else 14, 0), pady=6)
            self.lbl(f, naslov, 9, "bold", fg=INK2).pack(anchor="w")
            v = tk.StringVar(value=vrednost)
            w = (ttk.Combobox(f, textvariable=v, values=vals, state=state, width=width, font=self.f(10))
                 if vals is not None else ttk.Entry(f, textvariable=v, width=width, font=self.f(10)))
            w.pack(fill="x", pady=(4, 0))
            vars_[key] = v
            return w, f

        z = zapis or preset or {}
        w_grlo, _ = polje(1, 0, "Animal *", "grlo", z.get("grlo") or x or "", [g["id"] for g in sva])
        polje(1, 1, "Date *", "datum", z.get("datum") or date.today().strftime("%d/%m/%Y"))
        fd = tk.Frame(b, bg=SURF)
        fd.grid(row=2, column=0, columnspan=2, sticky="w", pady=6)
        self.lbl(fd, "Event * – click a type", 9, "bold", fg=INK2).pack(anchor="w")
        vars_["dogadjaj"] = tk.StringVar(value=z.get("dogadjaj", ""))
        self.izbor_dogadjaja(fd, vars_["dogadjaj"], po_redu=4).pack(anchor="w", pady=(4, 0))
        w_dog = fd
        w_vr, f_vr = polje(3, 0, "Value", "vrednost", z.get("vrednost", ""))
        w_par, f_par = polje(3, 1, "Partner (bull / ram…)", "partner", z.get("partner", ""), [])
        f_kar = self.okvir_karence(b, vars_, (z.get("karenca_meso", ""), z.get("karenca_mleko", "")))
        self.lbl(b, "Description / note", 9, "bold", fg=INK2).grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 4))
        opis = tk.Text(b, width=58, height=4, font=self.f(10), relief="flat", highlightthickness=1,
                       highlightbackground=LINE, highlightcolor=BRAND, wrap="word", padx=8, pady=6)
        opis.grid(row=5, column=0, columnspan=2, sticky="ew")
        opis.insert("1.0", z.get("opis", ""))
        info = self.lbl(b, "", 9, fg=INK2, wraplength=520, justify="left")
        info.grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 0))
        vz = tk.BooleanVar(value=True)
        chk = ttk.Checkbutton(b, text="Mark the animal as out of herd", variable=vz, style="Card.TCheckbutton")
        err = self.lbl(b, "", 10, "bold", fg=BAD)
        err.grid(row=8, column=0, columnspan=2, sticky="w", pady=(6, 0))
        nazivi_vr = {"Weighing": "Weight (kg)", "Vaccination": "Vaccine / product",
                     "Treatment": "Drug / diagnosis", "Sale": "Buyer / price", "Move": "New herd / location",
                     "Birth": "Number of young", "Mating / insemination": "Method (natural / AI)"}

        def prilagodi(*_):
            d = vars_["dogadjaj"].get()
            g = m.nadji(vars_["grlo"].get())
            if d in ("Treatment", "Vaccination"):
                f_par.grid_remove()
                f_kar.grid(row=3, column=1, sticky="nw", padx=(14, 0), pady=6)
            else:
                f_kar.grid_remove()
                f_par.grid()
            for w_ in f_vr.winfo_children():
                if isinstance(w_, tk.Label):
                    w_.config(text=nazivi_vr.get(d, "Value"))
            if g:
                a_ = m.A[g]
                suprotni = [k["id"] for k in m.grla() if k["vrsta"] == a_["vrsta"] and k["pol"] and k["pol"] != a_["pol"]]
                w_par.configure(values=suprotni)
            if d in ("Sale", "Death") and not zapis:
                chk.grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))
            else:
                chk.grid_remove()
            # ocekivani inbriding za parenje
            p_ = m.nadji(vars_["partner"].get())
            if d == "Mating / insemination" and g and p_ and g != p_:
                f = m.kinship(g, p_)
                t_, boja = procena(f, self.prag)
                info.config(text=f"Expected inbreeding of offspring: {pct(f)} ({t_}), shared genes {pct(m.r(g, p_))}.",
                            fg=boja)
            else:
                info.config(text="")
        for k in ("dogadjaj", "grlo", "partner"):
            vars_[k].trace_add("write", prilagodi)
        prilagodi()

        def sacuvaj(ev=None):
            g = m.nadji(vars_["grlo"].get())
            if not g or m.A[g]["osnivac"]:
                err.config(text="Choose an animal from the list (tag or ear tag).")
                return
            d = parsiraj_datum(vars_["datum"].get())
            if not d:
                err.config(text="The date is not valid. Enter it as 15/03/2026.")
                return
            dog = vars_["dogadjaj"].get().strip()
            if not dog:
                err.config(text="Choose an event.")
                return
            par = vars_["partner"].get().strip()
            par = m.nadji(par) or par
            leci = dog in ("Treatment", "Vaccination")
            podaci = dict(grlo=g, datum=d.strftime("%d/%m/%Y"), dogadjaj=dog, opis=opis.get("1.0", "end").strip(),
                          partner="" if leci else par, vrednost=vars_["vrednost"].get().strip(),
                          karenca_meso=self.broj_dana(vars_["karenca_meso"].get()) if leci else "",
                          karenca_mleko=self.broj_dana(vars_["karenca_mleko"].get()) if leci else "")
            if zapis:
                zapis.update(podaci)
            else:
                self.posle_novog_zapisa(self.ist.dodaj(**podaci))
                if dog in ("Sale", "Death") and vz.get():
                    r = self.red_za(g)
                    if r and r.get("status") == "aktivno":
                        r["status"] = "neaktivno"
                        self.ist.dodaj(g, AUTO, f"Status: out of herd ({dog.lower()})", automatski=True)
            win.grab_release()
            win.destroy()
            self.promena(f"Recorded in history: {g} – {dog}")
            if posle:
                try:
                    posle()
                except tk.TclError:
                    pass
            if getattr(self, "ist_win", None) and self.ist_win.winfo_exists() and hasattr(self.ist_win, "osvezi"):
                try:
                    self.ist_win.osvezi()
                except tk.TclError:
                    pass

        fb = tk.Frame(b, bg=SURF)
        fb.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        self.btn(fb, "Save", sacuvaj, "Primary").pack(side="right")
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=8)
        win.bind("<Escape>", lambda e: win.destroy())
        win.update_idletasks()
        r = parent or self.root
        win.geometry(f"+{r.winfo_rootx() + max(0, (r.winfo_width() - win.winfo_width()) // 2)}"
                     f"+{r.winfo_rooty() + max(20, (r.winfo_height() - win.winfo_height()) // 3)}")
        (w_dog if (x or zapis) else w_grlo).focus_set()

    @staticmethod
    def naziv_gen(g):
        return {1: "parent", 2: "grandparent", 3: "great-grandparent", 4: "great-great-grandparent"}.get(
            g, f"generation {g}")

    def stablo(self, x, gen=4, tab="stablo"):
        m = self.model
        if x not in m.A:
            return
        if getattr(self, "stablo_win", None) and self.stablo_win.winfo_exists():
            win = self.stablo_win
            for w in win.winfo_children():
                w.destroy()
        else:
            win = tk.Toplevel(self.root, bg=BG)
            win.geometry(f"1040x700+{self.root.winfo_rootx() + 30}+{self.root.winfo_rooty() + 20}")
            win.transient(self.root)
            win.bind("<Escape>", lambda e: win.destroy())
            self.stablo_win = win
        win.title(f"Pedigree of animal {x}")
        a = m.A[x]
        put = m.putanje(x)                       # SVI preci, sve generacije
        dubina = max((max(gs) for gs in put.values()), default=0)
        ponavlja = {y: gs for y, gs in put.items() if sum(gs.values()) > 1}
        gen = max(2, min(gen, min(7, dubina) + 1))

        # zaglavlje
        h = tk.Frame(win, bg=BG)
        h.pack(fill="x", padx=24, pady=(18, 6))
        self.cip(h, x, lambda: self.karton(x), big=True, bg=BG, siv=a["osnivac"]).pack(side="left")
        t = tk.Frame(h, bg=BG)
        t.pack(side="left", padx=14)
        self.lbl(t, f"Pedigree – {m.pol_naziv(x)}{', ' + a['rasa'] if a['rasa'] else ''}", 14, "bold").pack(anchor="w")
        por = a["poreklo"] + (f" – {a['poreklo_od']}" if a["poreklo_od"] else "") if a["poreklo"] else a["poreklo_od"]
        info = f"{len(put)} known ancestors in {dubina} generations   ·   inbreeding {pct(m.inbr(x))}"
        self.lbl(t, (f"{por}   ·   " if por else "") + info, 10, fg=INK2).pack(anchor="w")

        seg = tk.Frame(win, bg=BG)
        seg.pack(fill="x", padx=24, pady=(4, 8))
        sg = tk.Frame(seg, bg=SURF2)
        sg.pack(side="left")
        n_boc = len(self.bocni_srodnici(x))
        n_pot = sum(len(v) for v in self.potomci_po_kolenima(x).values())
        for k, tx in (("stablo", "Tree"), ("svi", f"All ancestors ({len(put)})"),
                      ("potomci", f"Descendants ({n_pot})"), ("bocni", f"Collateral relatives ({n_boc})")):
            akt = tab == k
            lb = tk.Label(sg, text=tx, font=self.f(10, "bold"), bg=SURF if akt else SURF2, fg=INK if akt else INK2,
                          padx=14, pady=6, cursor="hand2")
            lb.pack(side="left", padx=3, pady=3)
            lb.bind("<Button-1>", lambda e, k=k: self.stablo(x, gen, k))
        if tab == "stablo" and dubina >= 2:
            desno = tk.Frame(seg, bg=BG)
            desno.pack(side="right")
            self.lbl(desno, "Generations shown:", 9, "bold", fg=INK2).pack(side="left", padx=(0, 6))
            gv = tk.StringVar(value=str(gen - 1))
            cg = ttk.Combobox(desno, textvariable=gv, values=[str(i) for i in range(1, min(7, dubina) + 1)],
                              state="readonly", width=4, font=self.f(10))
            cg.pack(side="left")
            cg.bind("<<ComboboxSelected>>", lambda e: self.stablo(x, int(gv.get()) + 1, "stablo"))

        # preci koji se ponavljaju – kroz CEO rodoslov, ne samo prikazane generacije
        leg = tk.Frame(win, bg=BG)
        leg.pack(fill="x", padx=24, pady=(0, 6))
        if not put:
            self.lbl(leg, "No parents are entered for this animal.", 10, fg=INK2).pack(anchor="w")
        elif ponavlja:
            opis = ", ".join(f"{y} ({sum(gs.values())}×, gen. {' and '.join(str(g) for g in sorted(gs))})"
                             for y, gs in sorted(ponavlja.items(), key=lambda t: (min(t[1]), sort_key(t[0]))))
            self.lbl(leg, "Ancestors repeated in the whole pedigree (cause of inbreeding): " + opis, 9, fg=BAD,
                     wraplength=980, justify="left").pack(anchor="w")
            dalji = [y for y, gs in ponavlja.items() if min(gs) >= gen]
            if tab == "stablo" and dalji:
                self.lbl(leg, f"Some of them are beyond the generations shown ({', '.join(sorted(dalji, key=sort_key))})"
                              " – increase the number of generations or open the \"All ancestors\" tab.", 9, fg=WARN,
                         wraplength=980, justify="left").pack(anchor="w")
        else:
            self.lbl(leg, "No ancestor is repeated in the whole known pedigree.", 9, fg=OK).pack(anchor="w")

        if tab == "svi":
            self.svi_preci(win, x, put, ponavlja, dubina)
        elif tab == "bocni":
            self.tab_bocni(win, x)
        elif tab == "potomci":
            self.tab_potomci(win, x)
        else:
            self.crtaj_stablo(win, x, gen, ponavlja)
        win.lift()

    def potomci_po_kolenima(self, x):
        """{koleno: [potomak, ...]} – 1 deca, 2 unuci, 3 praunuci…"""
        m = self.model
        deca = {}
        for k, (o, mj) in m.P.items():
            for p in (o, mj):
                if p:
                    deca.setdefault(p, set()).add(k)
        rez, vidjeni, nivo, g = {}, {x}, {x}, 1
        while nivo and g <= 10:
            sl = set()
            for y in nivo:
                sl |= deca.get(y, set())
            sl -= vidjeni
            if not sl:
                break
            rez[g] = sorted(sl, key=lambda y: (m.A[y]["datum"] or date.max, sort_key(y)))
            vidjeni |= sl
            nivo, g = sl, g + 1
        return rez

    def tab_potomci(self, win, x):
        m = self.model
        po = self.potomci_po_kolenima(x)
        ukupno = sum(len(v) for v in po.values())
        legla = [z for z in self.cekaju_mindjusu(x)]
        nazivi_k = {1: "Children", 2: "Grandchildren", 3: "Great-grandchildren", 4: "Great-great-grandchildren"}
        opis = " · ".join(f"{nazivi_k.get(k, f'generation {k}')}: {len(v)}" for k, v in po.items())
        self.lbl(win, (f"In total {ukupno} descendants – {opis}." if ukupno else "No descendants recorded.") +
                 " Share of genes = how many genes the descendant carries from this animal.", 9, fg=INK2, wraplength=980,
                 justify="left").pack(anchor="w", padx=24, pady=(0, 6))
        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        if not ukupno and not legla:
            self.napomena(body, "This animal has no recorded descendants yet.", ok=True)
            return
        wrap, tv = self.tabela(body, [("id", "Offspring", 200), ("odnos", "Other parent", 200),
                                      ("pol", "Sex", 80), ("geni", "Share of genes", 95, "e"), ("rasa", "Breed", 120),
                                      ("rod", "Born", 95), ("st", "Status", 140)], visina=16)
        wrap.pack(fill="both", expand=True)
        tv.tag_configure("grupa", font=self.f(10, "bold"), background=SURF2)
        for k, lst in po.items():
            gr = tv.insert("", "end", text=f"  {nazivi_k.get(k, f'generation {k}')} ({len(lst)})", open=k <= 2,
                           tags=["grupa"])
            for y in lst:
                p = m.A[y]
                o, mj = m.P.get(y, (None, None))
                if k == 1:
                    drugi = mj if o == x else o
                    dr = f"{'dam' if o == x else 'sire'} {drugi}" if drugi else "unknown"
                else:
                    preko = [r for r in (o, mj) if r and (r in po.get(k - 1, []))]
                    dr = f"through {', '.join(preko)}" if preko else "–"
                st = ("deleted" if p.get("obrisano") else "in herd" if p["aktivno"] else "out of herd")
                tv.insert(gr, "end", iid=y, text="  " + y, image=self.img_tag, values=(
                    dr, m.pol_naziv(y), pct(m.doprinos(y, x)), p["rasa"] or "–",
                    fmt_datum(p["datum"]) or "–", st), tags=[] if p["aktivno"] else ["off"])
        if legla:
            gr = tv.insert("", "end", text="  Born, no ear tag yet", open=True, tags=["grupa"])
            for z in legla:
                uk, dd, vec = self.leglo(z)
                majka, otac = self.roditelji_parenja(z)
                tv.insert(gr, "end", text=f"  {uk - len(vec)} young", values=(
                    f"{majka or '?'} × {otac or '?'}", "–", "50,0%", "–", dd, "waiting for entry in the list"))
        tv.bind("<Double-1>", lambda e: self.otvori_iz(tv))
        self.lbl(body, "Double-click a descendant to open its card.", 9, fg=INK2).pack(anchor="w", pady=(8, 0))

    def bocni_srodnici(self, x):
        """Bočni (kolateralni) srodnici: braća/sestre, braća/sestre roditelja, nećaci, prvi rođaci.
        Vraća [(grupa, srodnik, opis odnosa)]."""
        m = self.model
        deca = {}
        for k, (o, mj) in m.P.items():
            for p in (o, mj):
                if p:
                    deca.setdefault(p, set()).add(k)

        def braca(y):
            o, mj = m.P.get(y, (None, None))
            rez = {}
            for s in (deca.get(o, set()) | deca.get(mj, set())) - {y}:
                so, sm = m.P.get(s, (None, None))
                if o and mj and so == o and sm == mj:
                    rez[s] = "pun"
                elif o and so == o:
                    rez[s] = "paternal"
                elif mj and sm == mj:
                    rez[s] = "maternal"
            return rez

        def naziv(s, z_, m_):
            return m_ if m.A.get(s, {}).get("pol") == "m" else z_ if m.A.get(s, {}).get("pol") == "z" else f"{m_} / {z_}"

        rez, vidjeni = [], {x}
        predci = m.preci(x)
        # 1) braća i sestre
        for s, tip in sorted(braca(x).items(), key=lambda t: (t[1] != "pun", sort_key(t[0]))):
            ime = naziv(s, "sister", "brother")
            rez.append(("Siblings", s, ("full " if tip == "pun" else "half-") + ime +
                        ("" if tip == "pun" else f" ({tip})")))
            vidjeni.add(s)
        # dvostruki odnos: brat/sestra koji je ujedno i dete brata/sestre
        sib = set(braca(x))
        for i, (g, y, op) in enumerate(rez):
            if any(p_ in sib for p_ in m.P.get(y, (None, None)) if p_):
                rez[i] = (g, y, op + f", also {naziv(y, 'niece', 'nephew')}")
        # 2) deca braće i sestara
        for s in list(braca(x)):
            for d in sorted(deca.get(s, set()) - vidjeni, key=sort_key):
                rez.append(("Nephews and nieces (children of siblings)", d, f"child of {s}"))
                vidjeni.add(d)
        # 3) braća i sestre roditelja
        o, mj = m.P.get(x, (None, None))
        strane = []
        for rod, strana, (mz, zz) in ((o, "paternal", ("uncle", "aunt")), (mj, "maternal", ("uncle", "aunt"))):
            if not rod:
                continue
            for s, tip in sorted(braca(rod).items(), key=lambda t: sort_key(t[0])):
                if s in vidjeni or s in predci:
                    continue
                rez.append(("Siblings of parents", s, f"{strana} {naziv(s, zz, mz)}" +
                            ("" if tip == "pun" else " (half)")))
                vidjeni.add(s)
                strane.append((s, strana))
        # 4) prvi rođaci (deca braće i sestara roditelja)
        for s, strana in strane:
            for d in sorted(deca.get(s, set()) - vidjeni, key=sort_key):
                if d in predci:
                    continue
                rez.append(("First cousins", d, f"child of {s} ({strana})"))
                vidjeni.add(d)
        return rez

    def tab_bocni(self, win, x):
        m = self.model
        sr = self.bocni_srodnici(x)
        self.lbl(win, "Collateral relatives are neither ancestors nor descendants, but share ancestors with this animal. "
                      "Be careful when mating – a high share of genes also means higher inbreeding of the young.", 9,
                 fg=INK2, wraplength=980, justify="left").pack(anchor="w", padx=24, pady=(0, 6))
        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        if not sr:
            self.napomena(body, "No known collateral relatives (siblings, cousins…) according to the recorded pedigree.",
                          ok=True)
            return
        wrap, tv = self.tabela(body, [("id", "Relative", 240), ("odnos", "Relation", 260), ("geni", "Shared genes", 120, "e"),
                                      ("f", "Inbreeding of young", 120, "e"), ("rasa", "Breed", 110),
                                      ("rod", "Born", 70), ("st", "Status", 140)], visina=16)
        wrap.pack(fill="both", expand=True)
        tv.tag_configure("grupa", font=self.f(10, "bold"), background=SURF2)
        grupe = {}
        for g, y, opis in sr:
            if g not in grupe:
                grupe[g] = tv.insert("", "end", text="  " + g, open=True, tags=["grupa"])
        brojac = {}
        for g, y, opis in sr:
            brojac[g] = brojac.get(g, 0) + 1
            p = m.A.get(y, {})
            r = m.r(x, y)
            suprotan = p.get("pol") and m.A[x]["pol"] and p.get("pol") != m.A[x]["pol"]
            f = m.kinship(x, y) if suprotan else None
            st = ("known only as a parent" if p.get("osnivac") else "deleted" if p.get("obrisano")
                  else "in herd" if p.get("aktivno") else "out of herd")
            tagovi = []
            if f is not None and f > self.prag and p.get("aktivno"):
                tagovi.append("bad")
            elif not p.get("aktivno"):
                tagovi.append("off")
            tv.insert(grupe[g], "end", iid=y, text="  " + y, image=self.img_tag_siv if p.get("osnivac") else self.img_tag,
                      values=(opis, pct(r), pct(f) if f is not None else "–", p.get("rasa") or "–",
                              p["datum"].year if p.get("datum") else "–", st), tags=tagovi)
        for g, iid in grupe.items():
            tv.item(iid, text=f"  {g} ({brojac[g]})")
        tv.bind("<Double-1>", lambda e: self.otvori_iz(tv))
        self.lbl(body, "Red = opposite sex in the herd with which the young would have inbreeding above the limit. "
                       "Double-click opens the card.", 9, fg=INK2).pack(anchor="w", pady=(8, 0))

    def svi_preci(self, win, x, put, ponavlja, dubina):
        m = self.model
        kom = m.kompletnost(x, dubina)
        if kom:
            self.lbl(win, "Pedigree completeness:  " + "   ".join(
                f"gen. {g}: {n}/{od}" for g, n, od in kom), 9, fg=INK2, wraplength=980, justify="left").pack(
                anchor="w", padx=24, pady=(0, 6))
        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=(0, 20))
        if not put:
            return
        wrap, tv = self.tabela(body, [("id", "Ancestor", 110), ("gen", "Generation", 220), ("put", "Times", 60, "e"),
                                      ("udeo", "Share of genes", 95, "e"), ("pol", "Sex", 75), ("mk", "Ear tag", 115),
                                      ("rasa", "Breed", 95), ("st", "Status", 160)], visina=16)
        wrap.pack(fill="both", expand=True)
        redovi = sorted(put.items(), key=lambda t: (-m.doprinos(x, t[0]), min(t[1]), sort_key(t[0])))
        for i, (y, gs) in enumerate(redovi):
            p = m.A[y]
            n = sum(gs.values())
            gen_txt = (f"{min(gs)}.  {self.naziv_gen(min(gs))}" if len(gs) == 1
                       else ", ".join(str(g) for g in sorted(gs)) + " generations")
            st = ("known only as a parent" if p["osnivac"] else "deleted" if p.get("obrisano")
                  else "in herd" if p["aktivno"] else "out of herd")
            tags = (["odd"] if i % 2 else []) + (["bad"] if y in ponavlja else []) + (["off"] if p["osnivac"] else [])
            tv.insert("", "end", iid=y, text="  " + y, image=self.img_tag_siv if p["osnivac"] else self.img_tag,
                      values=(gen_txt, f"{n}×", pct(m.doprinos(x, y)), m.pol_naziv(y), p["markica"] or "–",
                              p["rasa"] or "–", st), tags=tags)
        tv.bind("<Double-1>", lambda e: self.otvori_iz(tv))
        self.lbl(body, "Red = the ancestor appears more than once.  Double-click opens the ancestor's card.", 9, fg=INK2).pack(
            anchor="w", pady=(8, 0))

    def crtaj_stablo(self, win, x, gen, ponavlja):
        m = self.model
        sloti = {}
        def skupi(y, g, i):
            if g >= gen:
                return
            sloti[(g, i)] = y
            o, mj = m.P.get(y, (None, None)) if y else (None, None)
            skupi(o, g + 1, 2 * i)
            skupi(mj, g + 1, 2 * i + 1)
        skupi(x, 0, 0)
        self.lbl(win, "Each ancestor shows what share of genes the animal carries from it (sum of all paths). "
                      "Click an ancestor to show its pedigree, double-click opens its card.", 9, fg=INK2,
                 wraplength=980, justify="left").pack(anchor="w", padx=24)

        wrap = tk.Frame(win, bg=LINE)
        wrap.pack(fill="both", expand=True, padx=24, pady=(4, 20))
        c = tk.Canvas(wrap, bg=SURF, highlightthickness=0)
        sy = ttk.Scrollbar(wrap, orient="vertical", command=c.yview)
        sx = ttk.Scrollbar(wrap, orient="horizontal", command=c.xview)
        c.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        c.pack(side="left", fill="both", expand=True, padx=1, pady=1)

        BW, BH, GAP = 196, 50, 34
        listova = 2 ** (gen - 1)
        visina = max(480, listova * (BH + 12) + 40)
        for g in range(gen):
            nz = "Animal" if g == 0 else {1: "Parents", 2: "Grandparents", 3: "Great-grandparents",
                                         4: "Great-great-grandparents"}.get(g, f"generation {g}")
            c.create_text(20 + g * (BW + GAP) + BW / 2, 16, text=nz, font=self.f(9, "bold"), fill=INK2)
        pos = {}
        for (g, i) in sloti:
            pos[(g, i)] = (20 + g * (BW + GAP), 30 + (i + 0.5) * (visina - 40) / (2 ** g))
        for (g, i), (cx, cy) in pos.items():
            for k in (0, 1):
                ch = (g + 1, 2 * i + k)
                if ch in pos:
                    px, py = pos[ch]
                    mx = cx + BW + GAP / 2
                    c.create_line(cx + BW, cy, mx, cy, mx, py, px, py, fill=LINE, width=2)
        for (g, i), y in sloti.items():
            cx, cy = pos[(g, i)]
            x1, y1, x2, y2 = cx, cy - BH / 2, cx + BW, cy + BH / 2
            uloga = "" if g == 0 else ("sire" if i % 2 == 0 else "dam")
            if not y:
                c.create_polygon(self.zaobljeno(x1, y1, x2, y2, 8), smooth=True, fill=BG, outline="#C3CBC4", dash=(4, 3))
                c.create_text(x1 + 14, cy, text=f"{uloga} unknown" if uloga else "",
                              anchor="w", font=self.f(9), fill="#98A39B")
                continue
            p = m.A[y]
            rep_ = y in ponavlja
            tag = f"b{g}_{i}"
            c.create_polygon(self.zaobljeno(x1, y1, x2, y2, 8), smooth=True,
                             fill=SURF2 if p["osnivac"] else (BAD_BG if rep_ else SURF),
                             outline=BAD if rep_ else (BRAND if g == 0 else LINE), width=2 if (rep_ or g == 0) else 1,
                             tags=tag)
            c.create_rectangle(x1 + 1, y1 + 8, x1 + 6, y2 - 8, fill=TAG if not p["osnivac"] else "#C9D0CA",
                               outline="", tags=tag)
            c.create_text(x1 + 16, cy - 9, text=y, anchor="w", font=self.f(10, "bold"), fill=INK, tags=tag)
            if uloga:
                c.create_text(x2 - 10, cy - 9, text=f"{uloga} · {pct(m.doprinos(x, y))} genes", anchor="e",
                              font=self.f(8), fill=BAD if rep_ else INK2, tags=tag)
            sub = " · ".join(v for v in (p["markica"], p["rasa"]) if v)
            if p["datum"] and len(sub) + 7 <= 30:
                sub = (sub + " · " if sub else "") + str(p["datum"].year)
            sub = sub or ("known only as a parent" if p["osnivac"] else "")
            if len(sub) > 30:
                sub = sub[:29] + "…"
            c.create_text(x1 + 16, cy + 10, text=sub, anchor="w", font=self.f(8), fill=INK2, tags=tag)
            if g > 0:
                c.tag_bind(tag, "<Button-1>", lambda e, y=y: self.stablo(y, gen))
            c.tag_bind(tag, "<Double-1>", lambda e, y=y: self.karton(y))
            c.tag_bind(tag, "<Enter>", lambda e: c.config(cursor="hand2"))
            c.tag_bind(tag, "<Leave>", lambda e: c.config(cursor=""))
            # oznaka da ovaj predak ima jos predaka iza ivice prikaza
            if g == gen - 1 and any(m.P.get(y, (None, None))):
                c.create_text(x2 + 6, cy, text="›", anchor="w", font=self.f(14, "bold"), fill=INK2)
        c.configure(scrollregion=(0, 0, 40 + gen * (BW + GAP), visina + 10))
        c.bind("<MouseWheel>", lambda e: c.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units"))
        c.bind("<Button-4>", lambda e: c.yview_scroll(-1, "units"))
        c.bind("<Button-5>", lambda e: c.yview_scroll(1, "units"))

    # ---------- forma
    def forma(self, jid=None, preset=None):
        red = next((r for r in self.tab.redovi if r["id"] == jid), None) if jid else None
        posl = self.pod.get("poslednje", {})
        v = dict(red) if red else {"vrsta": posl.get("vrsta", "cattle"), "rasa": posl.get("rasa", ""),
                                   "stado": posl.get("stado", ""), "status": "aktivno", **(preset or {})}
        win = tk.Toplevel(self.root, bg=SURF)
        win.title("Edit animal" if red else "New animal")
        win.transient(self.root)
        win.resizable(False, False)
        win.grab_set()
        b = tk.Frame(win, bg=SURF)
        b.pack(fill="both", expand=True, padx=26, pady=22)
        self.lbl(b, "Edit animal" if red else "New animal", 17, "bold").grid(row=0, column=0, columnspan=2,
                                                                          sticky="w", pady=(0, 14))
        m = self.model
        sva = m.grla()
        vars_ = {}

        def polje(r, c, naslov, key, widget="entry", values=(), hint=""):
            f = tk.Frame(b, bg=SURF)
            f.grid(row=r, column=c, sticky="new", padx=(0 if c == 0 else 14, 0), pady=6)
            self.lbl(f, naslov, 9, "bold", fg=INK2).pack(anchor="w")
            var = tk.StringVar(value=v.get(key, ""))
            if widget == "entry":
                w = ttk.Entry(f, textvariable=var, width=28, font=self.f(10))
            else:
                w = ttk.Combobox(f, textvariable=var, values=list(values), width=26, font=self.f(10),
                                 state="readonly" if widget == "select" else "normal")
            w.pack(fill="x", pady=(4, 0))
            if hint:
                self.lbl(f, hint, 8, fg="#8A968D").pack(anchor="w")
            vars_[key] = var
            return w

        e_id = polje(1, 0, "Animal tag *", "id", hint="e.g. K12 or a name")
        polje(1, 1, "Ear tag", "markica", hint="e.g. UK 123456 700123")
        vr = polje(2, 0, "Species", "vrsta", "combo", VRSTE)
        polje(2, 1, "Breed", "rasa", "combo", sorted({a["rasa"] for a in sva if a["rasa"]}))

        pf = tk.Frame(b, bg=SURF)
        pf.grid(row=3, column=0, columnspan=2, sticky="w", pady=6)
        self.lbl(pf, "Sex *", 9, "bold", fg=INK2).pack(anchor="w")
        segf = tk.Frame(pf, bg=SURF2)
        segf.pack(anchor="w", pady=(4, 0))
        pol = tk.StringVar(value=v.get("pol", ""))
        seg_btns = {}

        def crtaj_seg(*_):
            n = nazivi(vrsta_norm(vars_["vrsta"].get()))
            for k, lb in seg_btns.items():
                akt = pol.get() == k
                lb.config(text=(n[1] if k == "z" else n[0]).capitalize(), bg=SURF if akt else SURF2,
                          fg=INK if akt else INK2)
        for k in ("z", "m"):
            lb = tk.Label(segf, font=self.f(10, "bold"), padx=20, pady=6, cursor="hand2")
            lb.pack(side="left", padx=3, pady=3)
            lb.bind("<Button-1>", lambda e, k=k: (pol.set(k), crtaj_seg()))
            seg_btns[k] = lb
        crtaj_seg()
        vars_["vrsta"].trace_add("write", crtaj_seg)

        d = parsiraj_datum(v.get("datum_rodjenja"))
        v["datum_rodjenja"] = d.strftime("%d/%m/%Y") if d else v.get("datum_rodjenja", "")
        polje(4, 0, "Date of birth", "datum_rodjenja", hint="day/month/year, e.g. 15/03/2022")
        polje(4, 1, "Herd / location", "stado", "combo", sorted({a["stado"] for a in sva if a["stado"]}))
        polje(5, 0, "Sire", "otac", "combo", sorted((a["id"] for a in m.A.values() if a["pol"] == "m"), key=sort_key),
              hint="tag or ear tag; empty if unknown")
        polje(5, 1, "Dam", "majka", "combo", sorted((a["id"] for a in m.A.values() if a["pol"] == "z"), key=sort_key),
              hint="tag or ear tag; empty if unknown")
        polje(6, 0, "Origin", "poreklo", "combo", POREKLA)
        polje(6, 1, "From whom / where", "poreklo_od", "combo",
              sorted({a["poreklo_od"] for a in sva if a["poreklo_od"]}))

        def auto_rasa(*_):
            """Kad su upisani otac i majka različitih rasa, rasa postaje 'hibrid O×M'."""
            o_, m_ = m.nadji(vars_["otac"].get()), m.nadji(vars_["majka"].get())
            if not o_ or not m_:
                return
            predlog = rasa_potomka(m.A[o_]["rasa"], m.A[m_]["rasa"])
            sad = vars_["rasa"].get().strip().lower()
            if predlog and (not sad or sad.startswith("hybrid") or sad == "crossbred"):
                vars_["rasa"].set(predlog)
        vars_["otac"].trace_add("write", auto_rasa)
        vars_["majka"].trace_add("write", auto_rasa)
        akt = tk.BooleanVar(value=v.get("status", "aktivno") != "neaktivno")
        ttk.Checkbutton(b, text="The animal is in the herd (not sold or dead)", variable=akt,
                        style="Card.TCheckbutton").grid(row=7, column=0, columnspan=2, sticky="w", pady=(10, 0))
        err = self.lbl(b, "", 10, "bold", fg=BAD)
        err.grid(row=8, column=0, columnspan=2, sticky="w", pady=(8, 0))

        def sacuvaj(dalje=False):
            nid = vars_["id"].get().strip()
            if not nid:
                err.config(text="Enter the animal's tag.")
                e_id.focus_set()
                return
            dup = next((r for r in self.tab.redovi if r["id"] == nid and r is not red), None)
            if dup:
                err.config(text=f"Animal {nid} already exists." + (" It is deleted – restore it from the \"Deleted animals\" list."
                                                             if dup.get("status") == "obrisano" else ""))
                return
            mk = vars_["markica"].get().strip()
            if mk:
                zauzeto = self.markica_zauzeta(mk, red)
                if zauzeto:
                    err.config(text=zauzeto)
                    return
            if pol.get() not in ("m", "z"):
                err.config(text="Choose the sex.")
                return
            dat = vars_["datum_rodjenja"].get().strip()
            if dat and not parsiraj_datum(dat):
                err.config(text="The date is not valid. Enter it as 15/03/2022.")
                return
            o, mj = vars_["otac"].get().strip(), vars_["majka"].get().strip()
            if nid in (o, mj):
                err.config(text="An animal cannot be its own parent.")
                return
            novi = dict(id=nid, markica=mk, vrsta=vrsta_norm(vars_["vrsta"].get()) or "cattle",
                        rasa=norm_rasa(vars_["rasa"].get()), pol=pol.get(),
                        datum_rodjenja=parsiraj_datum(dat).strftime("%d/%m/%Y") if dat else "",
                        otac=o, majka=mj, status=("obrisano" if red and red.get("status") == "obrisano"
                                                  else "aktivno" if akt.get() else "neaktivno"),
                        stado=vars_["stado"].get().strip(), poreklo=vars_["poreklo"].get().strip(),
                        poreklo_od=vars_["poreklo_od"].get().strip())
            if red:
                staro = red["id"]
                if staro != nid:
                    self.ist.preimenuj(staro, nid)
                    self.pods.preimenuj(staro, nid)
                    self.ist.dodaj(nid, AUTO, f"Tag changed: {staro} → {nid}", automatski=True)
                red["id"] = nid
                if norm(red.get("markica")) != norm(mk):
                    self.primeni_markicu(red, mk, zapamti=True)
                nazivi_polja = dict(vrsta="vrsta", rasa="rasa", pol="pol", datum_rodjenja="date of birth",
                                    otac="sire", majka="dam", status="status", stado="herd",
                                    poreklo="poreklo", poreklo_od="origin from")
                promene = [f"{n}: {red.get(k) or '–'} → {novi.get(k) or '–'}" for k, n in nazivi_polja.items()
                           if (red.get(k) or "") != (novi.get(k) or "")]
                if promene:
                    self.ist.dodaj(nid, AUTO, "Details changed – " + "; ".join(promene), automatski=True)
                red.update(novi)
                if staro != nid:
                    for r in self.tab.redovi:
                        if r.get("otac") == staro:
                            r["otac"] = nid
                        if r.get("majka") == staro:
                            r["majka"] = nid
            else:
                novi["uneto"] = sada()
                self.tab.redovi.append(novi)
                self.ist.dodaj(nid, AUTO, "Animal added to the records" +
                               (f" ({novi['poreklo']}{' – ' + novi['poreklo_od'] if novi['poreklo_od'] else ''})"
                                if novi["poreklo"] else ""), automatski=True)
                if dat:     # rodjenje se upisuje i u istoriju roditelja
                    for rod, drugi in ((mj, o), (o, mj)):
                        rid = self.model.nadji(rod) if rod else None
                        if rid and not self.model.A[rid]["osnivac"]:
                            self.ist.dodaj(rid, "Birth" if rod == mj else "Offspring",
                                           f"Offspring born {nid} ({'male' if pol.get() == 'm' else 'female'})",
                                           datum=parsiraj_datum(dat).strftime("%d/%m/%Y"),
                                           partner=self.model.nadji(drugi) or drugi or "", automatski=True)
            self.pod["poslednje"] = dict(vrsta=novi["vrsta"], rasa=novi["rasa"], stado=novi["stado"])
            self.sacuvaj_podesavanja()
            win.grab_release()
            win.destroy()
            self.promena(f"Saved: {nid}")
            if dalje:
                self.forma()

        def obrisi():
            if self.obrisi_grla([red["id"]], parent=win):
                win.destroy()

        fb = tk.Frame(win, bg=SURF)
        fb.pack(fill="x", padx=26, pady=(0, 20))
        tk.Frame(win, bg=LINE, height=1).place(in_=fb, relx=0, y=-10, relwidth=1)
        if red and red.get("status") != "obrisano":
            self.btn(fb, "Delete animal", obrisi, "Danger").pack(side="left")
        self.btn(fb, "Save", sacuvaj, "Primary").pack(side="right")
        if not red:
            self.btn(fb, "Save and add another", lambda: sacuvaj(True)).pack(side="right", padx=8)
        self.btn(fb, "Cancel", win.destroy).pack(side="right", padx=(0, 8) if red else 0)
        win.bind("<Return>", lambda e: sacuvaj())
        win.bind("<Escape>", lambda e: win.destroy())
        win.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_rooty() + max(20, (self.root.winfo_height() - win.winfo_height()) // 2)
        win.geometry(f"+{x}+{y}")
        (e_id if not v.get("id") else vr).focus_set()


def set_win_icon(window, ico):
    """Load the .ico and set it as the small and big window icon (title bar, Alt+Tab)."""
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        u.LoadImageW.restype = wintypes.HANDLE
        u.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT, ctypes.c_int, ctypes.c_int, wintypes.UINT]
        u.SendMessageW.restype = ctypes.c_ssize_t
        u.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_ssize_t]
        u.GetParent.restype = wintypes.HWND
        u.GetParent.argtypes = [wintypes.HWND]
        window.update_idletasks()
        hwnd = u.GetParent(window.winfo_id()) or window.winfo_id()
        for kind, metric, fallback in ((0, 49, 16), (1, 11, 32)):   # ICON_SMALL / ICON_BIG
            size = u.GetSystemMetrics(metric) or fallback
            h = u.LoadImageW(None, ico, 1, size, size, 0x10)       # IMAGE_ICON, LR_LOADFROMFILE
            if h:
                u.SendMessageW(hwnd, 0x0080, kind, h)              # WM_SETICON
    except Exception:
        pass


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        try:   # da Windows na traci zadataka prikaže logo programa, a ne Python
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Farmer.English")
        except Exception:
            pass
    sys.setrecursionlimit(20000)
    root = tk.Tk()

    def greska(exc, val, tb):
        messagebox.showerror("Error", "An error occurred:\n\n" + "".join(traceback.format_exception_only(exc, val)))
    root.report_callback_exception = greska
    App(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror("Farmer – error", traceback.format_exc())
        except Exception:
            print(traceback.format_exc())
