def round_format_size(s):
    from collections import defaultdict
    TB=1024*1024*1024*1024
    GB=1024*1024*1024
    MB=1024*1024
    KB=1024
    B=1
        
    _tp={TB:"Tb", GB:"Gb", MB:"Mb", KB:"Kb", B:"b"}
    print(_tp)
    
    if s >= TB:
        num = "%d%s" % (s // TB, _tp[TB])
    elif s >= GB:
        num = "%d%s" % (s // GB, _tp[GB])
    elif s >= MB:
        num = "%d%s" % (s // MB, _tp[MB])
    elif s >= KB:
        num = "%d%s" % (s // KB, _tp[KB])
    else: 
        num = "%f%s" % (s, _tp[B])
    return num
            
            
            