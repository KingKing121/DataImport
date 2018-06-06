from django.shortcuts import render_to_response


def student(request):
    list = [{id: 1, 'name': 'Jack'}, {id: 2, 'name': 'Rose'}]
    return render_to_response('student.html', {'students': list})