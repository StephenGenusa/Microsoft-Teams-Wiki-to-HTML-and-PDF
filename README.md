# Microsoft Teams Wiki to HTML and PDF

After waiting four years for Microsoft to realize that putting documentation into a Wiki *might* mean you are going to want to get that documentation out in some other format for distribution, say a PDF file, I have finally given up and written a parser for the grotesque MHT files generated by the Wiki editor.

The request for an export is one of the top ranked user requests, yet lower priority items have been worked on and completed, while Microsoft shows no signs of adding this. I am sure they eventually will, but this utility meets my immediate needs.

When I did my first conversion to HTML and then PDF I realized why, perhaps, Microsoft hasn't spent time creating an export. My first conversion, simply stripping the MHT header and dumping the HTML, resulted in a very  ugly HTML and PDF conversion. The Microsoft Wiki editor hides the grotesque nested font-family and font-size span elements and fudges the visual display so that users don't run from the Teams Wiki editor. I decided to delete the font-family and font-size tags, but had to do this in a repeated loop to get some pages presentable.

To use this utility,

1. Go to your SharePoint page where the wiki is located
2. Go to the Pages link
3. Download the Wiki directory. On my site it gave me a download named General.zip
4. Unzip the pages to a new directory. 
5. Run this command against the directory.
 
The pages from Sharepoint named "[title] - ##.mht" the utility reverses that so that the outputs are ## - [title].html/pdf This allows you to concat the PDFs in the order the pages were created, if desired which was more desirable in my case than a big dump of names, sortable only by title.

##Usage
<pre><code>
usage: teams\_wiki\_to\_html\_pdf.py [-h] [--wiki-input-path]
       teams\_wiki\_to\_html\_pdf.py --wiki-input-path "C:\Users\Username\Desktop\Your Path"
</pre>

