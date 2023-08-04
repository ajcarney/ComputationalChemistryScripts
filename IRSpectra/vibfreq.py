# GaussSum (http://gausssum.sf.net)
# Copyright (C) 2006-2013 Noel O'Boyle <baoilleach@gmail.com>
#
# This program is free software; you can redistribute and/or modify it
# under the terms of the GNU General Public License as published by the
# Free Software Foundation; either version 2, or (at your option) any later
# version.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY, without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# General Public License for more details.

import os
import math
import numpy
from cclib.parser import ccopen


class Spectrum(object):
    """Convolutes and stores spectrum data.

    Usage:
     Spectrum(start,end,numpts,peaks,width,formula)

    where
     peaks is [(pos,height),...]
     formula is a function such as gaussianpeak or delta
    

    >>> t = Spectrum(0,50,11,[[(10,1),(30,0.9),(35,1)]],5,delta)
    >>> t.spectrum
    array([[ 0.        ],
           [ 1.        ],
           [ 1.        ],
           [ 1.        ],
           [ 0.        ],
           [ 0.89999998],
           [ 1.89999998],
           [ 1.89999998],
           [ 1.        ],
           [ 0.        ],
           [ 0.        ]],'d')
    """
    def __init__(self,start,end,numpts,peaks,width,formula):
        self.start = start
        self.end = end
        self.numpts = numpts
        self.peaks = peaks
        self.width = width
        self.formula = formula

        # len(peaks) is the number of spectra in this object
        self.spectrum = numpy.zeros( (numpts,len(peaks)),"d")
        self.xvalues = numpy.arange(numpts)*float(end-start)/(numpts-1) + start
        for i in range(numpts):
            x = self.xvalues[i]
            for spectrumno in range(len(peaks)):
                for (pos,height) in peaks[spectrumno]:
                    self.spectrum[i,spectrumno] = self.spectrum[i,spectrumno] + formula(x,pos,height,width)


def lorentzian(x,peak,height,width):
    """The lorentzian curve.

    f(x) = a/(1+a)

    where a is FWHM**2/4
    """
    a = width**2./4.
    return float(height)*a/( (peak-x)**2 + a )

    
def irSpectra(inputFileName, outputFileName, start, end, numpts, FWHM, scaleFunction):
    print("Parsing file with cclib")
    ccData = ccopen(inputFileName).parse()

    freq = ccData.vibfreqs
    unscaledFreq = freq.copy()
    act = ccData.vibirs
    vibsyms = ccData.vibsyms
    scale = []
    
    for i in range(len(freq)):
        scalingFactor = scaleFunction(freq[i])
        freq[i] *= scalingFactor
        scale.append(scalingFactor)
    
    print("Convolving spectrum")
    spectrum = Spectrum(start,end,numpts, [list(zip(freq, act))], FWHM, lorentzian)
    
    print("Writing scaled spectrum to", outputFileName) 
    with open(outputFileName, "w") as outputFile:
        outputFile.write("Spectrum\t\t\tNormal Modes\n")
        outputFile.write("Freq (cm-1)\tIR act\t\tMode\tLabel\tFreq (cm-1)\tIR act\t")
        outputFile.write("Scaling factors\tUnscaled freq\n")
        
        width = end - start
        for x in range(numpts):
            if spectrum.spectrum[x,0]<1e-20:
                spectrum.spectrum[x,0] = 0.
            realx = width * (x + 1) / numpts + start
            outputFile.write(str(realx) + "\t" + str(spectrum.spectrum[x,0]))

            if x < len(freq): # Write the activities (assumes more pts to plot than freqs - fix this)
                outputFile.write("\t\t"+str(x+1)+"\t"+vibsyms[x]+"\t"+str(freq[x])+"\t"+str(act[x]))
                outputFile.write("\t"+str(scale[x])+"\t" + str(unscaledFreq[x]))
                
            outputFile.write("\n")
            
