<!doctype html>
<html lang="{{ app()->getLocale() }}">
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <title>Laravel</title>
    </head>
    <body>
        <form action="{{ route('generate') }}" method="post" enctype="multipart/form-data">
            <p>Select your CSV file:</p>
            <input type="file" name="csv_file">
            <p>Select your Logo file:</p>
            <input type="file" name="logo">
            <input type="hidden" name="_token" value="{{ csrf_token() }}">
            <br/>
            <br/>
            <input type="submit" name="submit" value="Generate">
        </form>
    </body>
</html>
