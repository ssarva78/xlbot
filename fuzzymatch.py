import difflib

def best_match(texts, text):

    if text is None or len(text) == 0:
        return None
                        
    normalize = lambda t: t.replace('_', '').\
                    replace('-','').\
                    replace(',','').\
                    replace(':','').\
                    replace("'",'').\
                    replace(' ', '').\
                    replace('\t', '').\
                    replace('%', 'percentage').\
                    lower()

    edit_distance = lambda t1, t2: len([d for d in \
                        difflib.ndiff(normalize(t1), normalize(t2)) \
                            if d[0] != ' '])

    candidates = [j for j in (i for i in \
                    ({t: edit_distance(t, text)} \
                        for t in texts if t is not None and len(t) > 0) \
                    if i.values()[0] < len(normalize(i.keys()[0])))]

  
    minimum_distance = min([i.values()[0] for i in candidates]) \
            if candidates is not None and len(candidates) > 0 else None
    
    return next(i for i in candidates \
                if i.values()[0] == minimum_distance).keys()[0] \
                    if minimum_distance is not None else None
    
    
