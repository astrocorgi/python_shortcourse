class Sequence():
    """Nucleic acid sequence data class"""
    def __init__(self,sequence='',quality=0):
        self.sequence=sequence
        self.quality=quality
    
    def __add__(self,other):
        result = Sequence()
        result.sequence = self.sequence + other.sequence
        result.quality = (len(self.sequence) * self.quality + len(other.sequence) * other.quality) / (len(self.sequence) + len(other.sequence))
        return result

    def __str__(self): #magic method, interacts with print function
        return 'Sequence: {}\nQuality: {}'.format(self.sequence,self.quality)

class DNASequence(Sequence):
    """DNA-specific sequence"""
    def __init__(self,sequence='',quality=0):
        print(sequence)
        if set(sequence) <= {'A','C','G','T'}:
            Sequence.__init__(self,sequence,quality)
        else:
            print("The input sequence isn't valid DNA input")

    def reverse_complement(self):
        rc_seq = ''
        complement = {'A':'T','T':'A','C':'G','G':'C'}
        for base in reversed(self.sequence):
            rc_seq = rc_seq + complement[base]
        return Sequence(rc_seq,self.quality)

    def transcribe(self):
        transcribed = ''.join([base if base !='T' else 'U' for base in self.sequence]) #for loop and if statement all in one line                                                                                                                          
        return RNASequence(transcribed,self.quality)

class RNASequence(Sequence):
    """RNA sequence"""
    def __init__(self,sequence='',quality=0):
        if set(sequence) <= {'A','C','G','U'}:
            Sequence.__init__(self,sequence,quality)
        else:
            raise Exception("The input sequence isn't valid RNA")
    
print('Making brain sample DNA class')
brain_sample = DNASequence('AGT',27)
print(brain_sample.transcribe())
brain_sample.reverse_complement()

#print('Creating liver sample RNA class...\n')
#liver_sample = RNASequence('AGU',30)

#print('Creating test sequence...\n')
#TestSequence = Sequence('ACGTC',14)

#print('Printing the addition of brain and liver sample')
#print(brain_sample + liver_sample)
