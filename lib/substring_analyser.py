from lib.ptrus_suffix_trees.STree import STree
import xlsxwriter
import re
from threading import Thread, Event

class SubstringAnalyser():
    '''Contains a dictionary with metadata about each text and analysis results populated by
the corresponding suffix tree. Output is via a generator method also contained in the dictionary.
Also has methods to save the data to an excel sheet at a user-specified filepath.'''

    def __init__(self, min_length=2, min_occurrences=2, spaced=False):
        '''Args:
spaced: whether the text has words split by spaces or not.
min_length: the minimum length in characters (in words if the text is spaced) of substrings in the results.
min_occurrences: minimum number of occurrences before a substring is included in the results.
Higher values for min_length and min_occurrences produce less results and slightly better performance.'''
        self.spaced = spaced
        self.min_length = min_length
        self.min_occurrences = min_occurrences
        self.data = []
        self.common = {'results' : [], 'clean_results' : []}
        self.common['output'] = self.get_output(self.common)

    def load(self, data_in):
        '''Data can be passed in as a string "text", a tuple (filename, text), or a list of tuples.'''
        if isinstance(data_in, str):
            data_in = [('', data_in)]
        if isinstance(data_in, tuple):
            data_in = [data_in]
        if isinstance(data_in, list):
            threads = []
            for i, d in enumerate(data_in):
                print('Loading {}'.format(i))
                self.data.append({'filename' : d[0], 'index' : i})
                threads.append( Thread(target=self.process_data, args=(d[1],i)) )
                threads[i].start()
            for t in threads:
                t.join()
        else:
            raise Exception('TermExtractor can only load strings or lists of strings.')

    def load_common(self):
        '''This method has to be called manually to populate the common substrings data.'''
        print('Loading common')
        if len(self.data) > 1:
            self.common['results'] = self.get_common(texts=list(d['text'] for d in self.data))

    def process_data(self, text, i):
        '''Splits spaced texts into a list of strings, then populates the dictionary entry for the text.'''
        if self.spaced:
            self.punctuation = '([{}]+)'.format(re.escape('\'!"()*,./:;<>?[]{} \n\t'))
            text = re.split(self.punctuation, text)
        else:
            self.punctuation = '([{}]+)'.format(re.escape(' \n\t。、（）「」　？・'))
        results = self.get_repeats(text)
        d = {'text': text, 'results' : results, 'clean_results' : []}
        d['output'] = self.get_output(d)
        self.data[i].update(d)

    def get_repeats(self, text):
        '''Uses a suffix tree to find all repeated substrings in the text.'''
        st = STree(text)
        
        def find_repeats(node):
            '''Recursive method to traverse the suffix tree.'''
            if node.is_leaf(): # Leaves never repeat
                return []
            repeats = []
            edge = st._edgeLabel(node, node.parent)
            if len(edge) >= 1: # Filters out single-letter strings and empty strings
                substring = st.word[node.idx:node.idx + node.depth]
                occurrences = len(node._get_leaves())
                length = len(list(s for s in substring if not re.search(self.punctuation, s)))
                if (occurrences >= self.min_occurrences) and (length >= self.min_length):
                    repeats.append((substring, occurrences))
            for (n,_) in node.transition_links:
                for s in find_repeats(n):
                    repeats.append(s)
            return repeats

        repeats = find_repeats(st.root)
        
        if self.spaced: # Converts spaced text back into a string.
            for i, repeat in enumerate(repeats):
                substring = ''.join(w for w in repeat[0])
                substring = substring.strip()
                repeats[i] = (substring, repeat[1])

        repeats.sort(key=lambda r: len(r[0])) #sort by length
        repeats.sort(key=lambda r: r[1]) #sort by number of occurrences

        return repeats

    def get_common(self, texts):
        '''Uses a generalised suffix tree to find all common substrings between the texts.'''
        gst = STree(texts, gst=True)
        
        def find_common(node):
            '''Recursive method to traverse the GST.'''
            nodes = []
            for (n,_) in node.transition_links:
                if len(n.generalized_idxs) > 1:
                    for c in find_common(n):
                        nodes.append(c)
            if nodes == []:
                return [node]
            return nodes

        common_nodes = find_common(gst.root)
        common_nodes.sort(key=lambda n: n.depth)
        common_substrings = []
        for node in common_nodes:
            substring = gst.word[node.idx:node.idx+node.depth]
            length = len(list(s for s in substring if not re.search(self.punctuation, s)))
            if length >= self.min_length:
                common_substrings.append((substring,node.generalized_idxs))

        for i, cs in enumerate(common_substrings): #GST uses list format, so both spaced and nonspaced text has to be converted back to strings.
            substring = ''.join(w for w in cs[0])
            substring = substring.strip()
            common_substrings[i] = (substring, cs[1])

        return common_substrings

    def get_output(self, data):
        '''Generator for output - I don't know how to implement an online ST, so redundant substrings are filtered
iteratively, which takes non-linear time. Generator gives a better UX until I learn how to improve performance.'''
        while True:
            try:
                result = data['results'].pop()
                if not any (result[0] in r[0] for r in data['clean_results']):
                    data['clean_results'].append(result)
                    yield result
            except IndexError:
                break

    def save_output(self, path):
        '''Manages threads for saving the output.'''
        wb = xlsxwriter.Workbook(path)

        threads = []
        if len(self.data) > 1:
            threads.append( Thread(target=self.save_common, args=(wb,)) )
        for d in self.data:
            threads.append(Thread(target=self.save_repeats, args=(d, wb)))
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        wb.close()

    def save_common(self, wb):
        '''Writes out results to excel in three-column format.'''
        sheet = wb.add_worksheet('Common substrings')
        sheet.set_column(0, 0, 60)
        sheet.write(0, 0, 'SUBSTRING')
        sheet.write(0, 1, 'APPEARS IN')
        sheet.write(0, 2, 'LENGTH')

        i = 1
        while True:
            try:
                out = next(self.common['output'])
                sheet.write(i, 0, out[0])
                sheet.write(i, 1, repr(out[1]).strip('{}'))
                sheet.write(i, 2, len(out[0]))
            except StopIteration:
                sheet.autofilter(0 ,0 ,i, 2)
                break
            i = i + 1

    def save_repeats(self, d, wb):
        '''Writes out results to excel in three-column format.'''
        excel_banned= '[{}]'.format(re.escape('[]:*?/\\'))
        filename = re.sub(excel_banned, '', d['filename'])
        sheet = wb.add_worksheet('{}： {}'.format(d['index'], d['filename'][:20]))
        sheet.set_column(0, 0, 60)
        sheet.write(0, 0, 'SUBSTRING')
        sheet.write(0, 1, 'OCCURRENCES')
        sheet.write(0, 2, 'LENGTH')
        
        i = 1
        while True:
            try:
                out = next(d['output'])
                sheet.write(i, 0, out[0])
                sheet.write(i, 1, out[1])
                sheet.write(i, 2, len(out[0]))
            except StopIteration:
                sheet.autofilter(0 ,0 ,i, 2)
                break
            i = i + 1
