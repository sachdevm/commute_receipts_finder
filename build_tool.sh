#!/bin/bash

cd URL2PDF
/usr/bin/xcodebuild -project DownloadURLsAsPDFs.xcodeproj
cp build/Release/Download\ URLs\ as\ PDFs.action/Contents/Resources/url2pdf ../urltopdf
cd ..
chmod +x url2pdf
