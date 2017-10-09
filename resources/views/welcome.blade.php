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
            {{-- <input type="button" name="docx" id="docx" value="Print"> --}}
        </form>
    </body>
</html>
{{-- <script src="https://code.jquery.com/jquery-3.2.1.min.js"></script>
<script src="assets/js/docxtemplater/docxtemplater-latest.min.js"></script>
<script src="assets/js/docxtemplater/docxtemplater-image-module-latest.min.js"></script>
<script src="assets/js/docxtemplater/jszip.min.js"></script>
<script src="assets/js/docxtemplater/file-saver.min.js"></script>
<script src="assets/js/docxtemplater/jszip-utils.js"></script>
<script type="text/javascript">
    function loadFile(url, callback) {
        JSZipUtils.getBinaryContent(url, callback);
    }

    function getImage(url, callback) {
        loadFile(url, function(error, content) {
            return callback(content);
        });
    }

    $('#docx').click(function(e) {
        e.preventDefault();
        loadFile("assets/template_js.docx", function(error, content) {
            if (error) throw error;
            var zip = new JSZip(content);
            var doc = new Docxtemplater().loadZip(zip);

            $.get("{{ route('csvdata') }}")
            .done(function(data, status) {
                getImage(data.logo, function(logo) {
                    var opts = {};
                    opts.centered = false;
                    opts.getImage = function(tagValue, tagName) {
                        return logo;
                    };
                    opts.getSize = function(img, tagValue, tagName) {
                        return [200, 100];
                    };
                    var imageModule = new window.ImageModule(opts);

                    doc.attachModule(imageModule);
                    doc.setData(data);
                    try {
                        doc.render();
                    }
                    catch (error) {
                        var e = {
                            message: error.message,
                            name: error.name,
                            stack: error.stack,
                            properties: error.properties,
                        };
                        console.log(JSON.stringify({error: e}));
                        // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
                        throw error;
                    }

                    var out = doc.getZip().generate({
                        type: "blob",
                        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    });
                    saveAs(out, "output.docx");
                });
            })
            .fail(function(data, status) {

            });
        });
    });
</script> --}}