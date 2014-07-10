class movie:

    def __init__(self,iden,title,resWidth,resHei,filepath,dual):
            self.iden=iden
            self.title=title
            self.reswidth=resWidth
            self.reshei=resHei
            self.filepath=filepath
            self.dual=dual 
    
    def __cmp__(self,other):
                if self.iden==other.iden:
                    return 1
                else:
                    return 0
    
    def __eq__(self,other):
            if self.iden==other.iden:
                return 1
            else:
                return 0