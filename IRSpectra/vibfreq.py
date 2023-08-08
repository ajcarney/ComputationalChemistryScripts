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

import numpy
from cclib.parser import ccopen


def broadenSpectrum(start, end, numpts, peaks, width, formula):
    """
    Broadens spectrum data. Creates a distribution function around
    each peak (pos, height) and adds up the contributions from each 
    distribution over numpts from start to end to create a spectrum
    that looks closer to one obtained by experiment

    
    formula is a function such as gaussianpeak or delta
    """
    spectrum = numpy.zeros(numpts,"d")
    xvalues = numpy.linspace(start, end, numpts)
    print(xvalues, start)
    for i in range(numpts):
        x = xvalues[i]
        for pos, height in peaks:
            spectrum[i] = spectrum[i] + formula(x, pos, height, width)

    return xvalues, spectrum



def lorentzian(x, peak, height, width):
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
    
    print("Broadening spectrum")
    xvalues, spectrum = broadenSpectrum(start, end, numpts, list(zip(freq, act)), FWHM, lorentzian)
    
    print("Writing scaled spectrum to", outputFileName) 
    with open(outputFileName, "w") as outputFile:
        outputFile.write("Spectrum\t\t\tNormal Modes\n")
        outputFile.write("Freq (cm-1)\tIR act\t\tMode\tLabel\tFreq (cm-1)\tIR act\t")
        outputFile.write("Scaling factors\tUnscaled freq\n")
        
        for i in range(numpts):
            if spectrum[i] < 1e-20:
                spectrum[i] = 0.

            outputFile.write(str(xvalues[i]) + "\t" + str(spectrum[i]))

            if i < len(freq): # Write the activities (assumes more pts to plot than freqs - fix this)
                outputFile.write("\t\t"+str(i+1)+"\t"+vibsyms[i]+"\t"+str(freq[i])+"\t"+str(act[i]))
                outputFile.write("\t"+str(scale[i])+"\t" + str(unscaledFreq[i]))
                
            outputFile.write("\n")
            
