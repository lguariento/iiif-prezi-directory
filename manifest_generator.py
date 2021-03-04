from iiif_prezi.factory import ManifestFactory
from openpyxl import load_workbook
import os
import warnings
warnings.simplefilter("ignore")

spreadsheet = load_workbook(filename="McLagan_MSGen1042.xlsx", read_only=True)
sheet = spreadsheet.active

ms_with_ranges =["MS Gen 1042/10","MS Gen 1042/102","MS Gen 1042/105","MS Gen 1042/106a","MS Gen 1042/106b","MS Gen 1042/109","MS Gen 1042/111","MS Gen 1042/114","MS Gen 1042/115","MS Gen 1042/118","MS Gen 1042/120","MS Gen 1042/122","MS Gen 1042/125","MS Gen 1042/126","MS Gen 1042/129a","MS Gen 1042/129b","MS Gen 1042/13","MS Gen 1042/130","MS Gen 1042/132","MS Gen 1042/135b","MS Gen 1042/136","MS Gen 1042/137","MS Gen 1042/139","MS Gen 1042/14","MS Gen 1042/140","MS Gen 1042/141","MS Gen 1042/142","MS Gen 1042/143","MS Gen 1042/145","MS Gen 1042/146","MS Gen 1042/148","MS Gen 1042/150","MS Gen 1042/151","MS Gen 1042/153","MS Gen 1042/154","MS Gen 1042/156","MS Gen 1042/160","MS Gen 1042/161","MS Gen 1042/162","MS Gen 1042/163","MS Gen 1042/165b","MS Gen 1042/166","MS Gen 1042/167","MS Gen 1042/168","MS Gen 1042/169","MS Gen 1042/170","MS Gen 1042/177","MS Gen 1042/18","MS Gen 1042/180","MS Gen 1042/181","MS Gen 1042/184","MS Gen 1042/185","MS Gen 1042/186","MS Gen 1042/187","MS Gen 1042/19","MS Gen 1042/190","MS Gen 1042/192","MS Gen 1042/193","MS Gen 1042/194","MS Gen 1042/195","MS Gen 1042/199","MS Gen 1042/2","MS Gen 1042/20","MS Gen 1042/200","MS Gen 1042/201","MS Gen 1042/204","MS Gen 1042/205","MS Gen 1042/209","MS Gen 1042/210","MS Gen 1042/211","MS Gen 1042/213","MS Gen 1042/216","MS Gen 1042/219","MS Gen 1042/22","MS Gen 1042/222a","MS Gen 1042/222b","MS Gen 1042/222c","MS Gen 1042/225","MS Gen 1042/226","MS Gen 1042/227","MS Gen 1042/229","MS Gen 1042/23","MS Gen 1042/230","MS Gen 1042/233","MS Gen 1042/235","MS Gen 1042/239","MS Gen 1042/240","MS Gen 1042/241","MS Gen 1042/244","MS Gen 1042/25","MS Gen 1042/26","MS Gen 1042/27","MS Gen 1042/29","MS Gen 1042/3","MS Gen 1042/33","MS Gen 1042/36","MS Gen 1042/39","MS Gen 1042/4","MS Gen 1042/45","MS Gen 1042/47","MS Gen 1042/5","MS Gen 1042/50","MS Gen 1042/51","MS Gen 1042/52","MS Gen 1042/53","MS Gen 1042/54","MS Gen 1042/58","MS Gen 1042/59","MS Gen 1042/61","MS Gen 1042/62","MS Gen 1042/64","MS Gen 1042/67","MS Gen 1042/68","MS Gen 1042/69","MS Gen 1042/70","MS Gen 1042/73","MS Gen 1042/76","MS Gen 1042/77","MS Gen 1042/80","MS Gen 1042/81","MS Gen 1042/82","MS Gen 1042/83","MS Gen 1042/85","MS Gen 1042/87","MS Gen 1042/89","MS Gen 1042/9","MS Gen 1042/90","MS Gen 1042/91","MS Gen 1042/94","MS Gen 1042/96","MS Gen 1042/97","MS Gen 1042/98","MS Gen 1042/99"]

prezi_dir = "/tmp"

fac = ManifestFactory()
fac.set_debug("error")
fac.set_iiif_image_info()
fac.set_base_prezi_uri("https://iiif.gla.ac.uk/")
fac.set_base_prezi_dir(prezi_dir)

#mflbl = image_dir.split('/')[7]

ranges_done = []
manifests = []

for ms in sheet.iter_rows(min_row=3, values_only=True):
    if (ms[2] == None) and (ms[0] not in ranges_done): # Check that it's not a range row.
        manifests.append(ms[0]) # Add it to the list of created manifests
        title = ms[0]
        manifest = fac.manifest(ident="manifest/" + ms[1] + ".json", label=title)
        manifest.set_metadata({"Physical description": ms[4], 'Bibliography': ms[8], 'Library record ID': str(ms[9])})
        manifest.attribution = "<span>Photo: © Archives & Special Collections, University of Glasgow Library. Terms of use: <a href=\"https://creativecommons.org/licenses/by-nc/4.0/\">CC-BY-NC</a></span>"
        if ms[7] != None:
            manifest.set_metadata({"Commentary": ms[7]})
        manifest.description = ms[6]
        seq = manifest.sequence(ident=ms[1] + ".json", label='Current page order')
        image_dir = "/Users/luca/Documents/development/iiif-prezi/images/" + ms[1] + "/images/"
        fac.set_base_image_dir(image_dir)
        fac.set_base_image_uri("https://iiif.gla.ac.uk/iiif/" + ms[1] + "/images")
        print('Getting data for ' + ms[0])
        for fn in sorted(os.listdir(image_dir)):

            ident = fn[:-4]
            cvs_title = ident.replace("_", " ").title().split(' ')[-1]
            cvs = seq.canvas(ident=ident, label=cvs_title)
            #cvs.add_image_annotation(fn, True)

            anno = cvs.annotation(ident=ident + ".json", label=title)
            image = anno.image(ident=fn, iiif=True)
            image.set_hw_from_iiif()
            cvs.set_hw(image.height, image.width)

        if ms[0] not in ms_with_ranges: # If it's not a ms with ranges add the title and save the manifest
            if ms[5] != None:
                manifest.set_metadata({"Title": ms[5]}) # Add the title of the poem if not null
            data = manifest.toString(compact=False)
            fh = open(ms[1] + '.json', mode="w", encoding="utf-8")
            fh.write(data)
            fh.close()
            print('Manifest for ' + ms[0] + ' saved')

        else: # ... else it has ranges defined, which we'll take care of.
            for poem in sheet.iter_rows(min_row=3, values_only=True): # Loops again through all the cells...
                if (poem[0] == ms[0]) and (poem[3] != None): # ... and find the ones belonging to the manuscript we're building the manifest for, and which have something in the 'range folios' cell.
                    print('found poem for ' + poem[1] + '(' + poem[2] + ')')
                    rng = manifest.range(ident=poem[1] + '/' + poem[2], label=poem[5])
                    if poem[6] != None:
                        rng.set_metadata({"Description": poem[6]})
                    if ',' in str(poem[3]): # Checks if there's more than one value, i.e. more than one page in this range.
                        for canvas in poem[3].split(','): # If so, loops through these comma-separated values.
                            rng.add_canvas(seq.canvases[int(canvas) - 1])
                            print('Added folio ' + canvas + ' to range ' + poem[5])
                    else: # If not add the only page that there is.
                        rng.add_canvas(seq.canvases[int(poem[3]) - 1])
                    #c = seq.canvas(ident="canvas1", label="Page 1")
                    #r.add_canvas(c)

                    ranges_done.append(ms[0]) # Append this manuscript to the ranges_done list so that we avoid being picked up by the first loop.

            data = manifest.toString(compact=False)
            fh = open(ms[1] + '.json', mode="w", encoding="utf-8")
            fh.write(data)
            fh.close()
            print('Manifest for ' + ms[0] + ' saved')


    elif ms[0] in ranges_done:
        pass




#manifest.toFile(compact=False)
