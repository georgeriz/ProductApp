import sqlite3

class WordChange:
    """transform, add, delete are supported
    entries stored at "data_files/word_change.txt"
    all entries must be in lower case"""
    def __init__(self):
        self.word_change = {}
        connection = sqlite3.connect("data_files/word_change.db")
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM synonyms")
        for row in cursor.fetchall():
            self.word_change[row[0]] = row[1]
        connection.close()

    def _find(self, a):
        if a in self.word_change:
            return self.word_change[a]
        return None

    def transform(self, input_term):
        """return the correct form of the input term, the same if not existent"""
        b = self._find(input_term)
        if b:
            return b
        else:
            return input_term

    def add(self, input_term, output_term):
        """add a new correction, returns False if input term already exists"""
        input_term = input_term.strip().lower()
        output_term = output_term.strip().lower()
        if len(input_term)==0 or len(output_term)==0 or self._find(input_term):
            return False
        #
        connection = sqlite3.connect("data_files/word_change.db")
        cursor = connection.cursor()
        cursor.execute("INSERT INTO synonyms VALUES (?,?)", (input_term, output_term))
        connection.commit()
        connection.close()
        #
        self.word_change[input_term] = output_term
        return True

    def delete(self, input_term):
        """delete entry for the input term, returns False if not existent"""
        input_term = input_term.strip().lower()
        if self._find(input_term):
            #
            connection = sqlite3.connect("data_files/word_change.db")
            cursor = connection.cursor()
            cursor.execute("DELETE FROM synonyms WHERE general=?", (input_term,))
            connection.commit()
            connection.close()
            #
            del self.word_change[input_term]
            return True
        else:
            return False