

@debug.get_steps()
#@debug.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('nit', doc='value to find similar in vector')
@xw.arg('outputs', doc='value to find similar in vector')
@xw.arg('no', int, doc='value to find similar in vector')
@xw.ret(expand='table', header=False, index=False, empty='NA')  #
def xf_search_web_nit(nit, outputs=None, no=None):
    """search the web for nit or id data and return found information"""
    raise NotImplementedError
    from common.cache import JsonDictFile
    from consulta_nit import AsyncNitAndIDFinder

    cache_nit = JsonDictFile("nits_and_ids.dict")
    with AsyncNitAndIDFinder(no=2, cache=cache_nit, automatic=True, spawn_in_main=False) as nit_finder_async:
        ids = ["91092908904", "13548629", "41477494", "890527", "41425269", "98011456898"]
        for id_ in ids:
            # only put if it is not in cache
            nit_finder_async.put(id_, update=False)

        for id_ in ids:
            time1 = time.time()
            data2 = nit_finder_async.request(id_)
            print(F"ID '{id_}' data: {data2}")
            print(F"time now:{time.time()-time1}")

        time2 = time.time()
    print(F"time closing:{time.time()-time2}")