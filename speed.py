# SPEED python module (created for use with HEEDS optimizaton toolset)
# James R. Dorris, Malcolm I. McGilp
# 7th July 2015
# Last edited 21st September 2015

# Requires the following free Python packages:
# pywin32 - required for the ActiveX connection to SPEED
# matplotlib - required for creation of matlab-like plots
# scipy - required for optimization toolkit used in advanced SPEED routines, useful for optimization and automation
# numpy - required by scipy
# PIL - required for generation of plot images
# *** All of the above are available in winPython - the Recommended approach for Python on Windows.

# base python modules
import sys, os, os.path, shutil, re, math, csv
# from pywin32 package
import win32com, win32com.client.dynamic, win32api, pythoncom, datetime, time
from win32com.test.util import CheckClean
from win32com.client import gencache
from win32com.shell import shell, shellcon
# from numpy package
import numpy as np
# from scipy package
import scipy.optimize as sciopt
import scipy.interpolate as interp
# from matplotlib package
import matplotlib.pyplot as plt
import matplotlib.cm as cm
# from PIL package
from PIL import Image


class speedWrapper:
	'speedWrapper Class creates an easy to use interface and helper function to interact with SPEED'

	def __init__(self, speed_file, dID=None, verbose=False):
		self.speed_file = speed_file
		#
		name, ext = os.path.splitext(speed_file)
		ext = ext.lower()

		dir = name
		name = dir.split("\\")[-1]
		dir = dir.replace(name, "")
		self.speed_fname = name
		self.speed_bname = name
		self.speed_ext = ext
		self.speed_dir = dir

		if dID != None:
			# it means we've asked to append dID to the filename and make a copy of it in the current working directory
			self.speed_fname = self.speed_fname+dID
			self.speed_dir = self.GetCWD()+"\\"
			self.speed_file = self.speed_dir+self.speed_fname+self.speed_ext
			try:
				shutil.copy(speed_file, self.speed_file)
				speed_file = self.speed_file
			except (IOError, WindowsError):
				self.Message("speedWrapper: Could not copy speed file to CWD with ID="+dID+" ("+speed_file+","+self.speed_file+")", err=1)

		self.torque_label = ['Nm', 'Nmm', 'kgm', 'kgcm', 'gcm', 'lbin', 'lbft', 'ozft', 'ozin']
		self.torque_factor = [1.0, 0.001, 1./9.80665, 100./9.80665, 100000./9.80665, 8.8507457676,	0.73756214728, 11.800994357, 141.61193228]
		self.flux_label = ['Wb', 'mWb', 'Maxwl', 'Lines', 'kLine']
		self.flux_factor = [1.0, 0.001, 1./100000000, 1./100000000, 1./100000]

		self.heeds_out = None

		if verbose != False:
			self.verbose = True
		else:
			self.verbose = False

		#Prepare speedWrapper log file and error handling
		dt = datetime.datetime.now()
		self.start_time = dt.strftime("%Y-%m-%d %H:%M:%S.")+ str(math.floor(dt.microsecond/10000))
		self.log_file_name = self.GetCWD()+"\\" + self.speed_fname+"_speedWrapper.log"
		try:
			self.log_file = open(self.log_file_name, "w")
			self.Message("speedWrapper started at "+self.start_time)
		except (IOError, WindowsError):
			raise IOError("Cannot Open Log File ("+self.log_file_name+")")

		if ext == '.bd4':
			self.appname = "wbdc32"
		elif ext == '.wfc':
			self.appname = "wwfc32"
		elif ext == ".im1":
			self.appname = "wimd32"
		elif ext == ".srd":
			self.appname = "wsrd32"
		elif ext == ".xsrd":
			self.appname = "wsrd32"
		elif ext == ".dcm":
			self.appname = "wdcm32"
		elif ext == ".axm":
			self.appname = "waxm32"
		else:
			self.Message("Improper SPEED file specified ("+speed_file+")", err=-1)

		try:
			self.wapp = win32com.client.dynamic.Dispatch(self.appname.upper()+".AppAutomation")
			self.wpi = self.updatePIparams(self.wapp, self.appname+"params")
			
			self.des = self.wapp.CreateNewDesign()
			
			res = self.des.LoadFromFile(self.speed_file)
			if res > 0:
				self.Message("Opening file ("+self.speed_file+"), returns ["+str(res)+"]: "+self.des.GetErrorMsg(res))
			if res < 0:
				self.Message("Opening file ("+self.speed_file+"), returns ["+str(res)+"]: "+self.des.GetErrorMsg(res), err=res)
				del self
		except pythoncom.com_error:
			self.Message("Failed to start "+self.appname, err=-1)
			del self

	def Message(self, msg, err=None):
		# writes output to the logfile and if verbose is set, it will also write to cmd line

		if err == None:
			if self.verbose:
				print(msg)
		elif err == 0:
			print(msg)
		elif err > 0:
			msg = "Warning: "+msg
			print(msg)
		elif err < 0:
			msg = "Error: "+msg
			print(msg)

		if not self.log_file.closed:
			self.log_file.write(msg+'\n')

		if err !=None and err < 0:
			raise RuntimeError(msg, err)

	
	def setParams(self, inParams):
		for key in inParams.keys():
			try:
				res = self.des.SetVariable(self.wpi['pi'+key.upper()], inParams[key])
			except KeyError:
				pID = self.des.getIndexFromName(key)
				res = self.des.SetVariable(pID, inParams[key])


	def getParam(self, outParam):
		try:
			out = self.des.GetVariable(self.wpi['pi'+outParam.upper()])
		except KeyError:
			pID = self.des.getIndexFromName(outParam)
			out = self.des.GetVariable(pID)
		return out

	def getParamName(self, param):
		try:
			out = self.des.GetVariableName(self.wpi['pi'+param.upper()])
		except KeyError:
			pID = self.des.getIndexFromName(param)
			out = self.des.GetVariableName(pID)
		return out

	def getWaveForm(self, waveVar):
		return self.des.GetWaveForm(waveVar)

	def Static(self):
		if self.speed_ext == '.bd4':
			return self.des.DoStaticSimulation()
		elif self.speed_ext == '.im1':
			return self.des.DoSteadyStateAnalysis()
		elif self.speed_ext == ".srd":
			return self.des.DoDynamicDesign()
		elif self.speed_ext == ".xsrd":
			return self.des.DoDynamicDesign()
		elif self.speed_ext == ".wfc":
			return self.des.DoStaticDesign()
		elif self.speed_ext == ".dcm":
			return self.des.DoDynamicDesign()
		elif self.speed_ext == ".axm":
			return self.des.DoStaticDesign()
		else:
			self.Message("Improper SPEED file specified, "+self.speed_ext, err=1)

	def Dynamic(self):
		if self.speed_ext == '.bd4':
			return self.des.DoDynamicSimulation()
		elif self.speed_ext == '.im1':
			return self.des.DoDriveSimulation()
		elif self.speed_ext == ".srd":
			return self.des.DoDynamicDesign()
		elif self.speed_ext == ".xsrd":
			return self.des.DoDynamicDesign()
		elif self.speed_ext == ".wfc":
			return self.des.DoStaticDesign()
		elif self.speed_ext == ".dcm":
			return self.des.DoDynamicDesign()
		elif self.speed_ext == ".axm":
			return self.des.DoDynamicDesign()
		else:
			self.Message("Improper SPEED file specified, "+self.speed_ext, err=1)

	def RunGoFER(self):
		bname = self.speed_dir+self.speed_fname
		gdf = win32com.client.dynamic.Dispatch("gdfed32.AppAutomation")

		if gdf.LoadFromFile(bname+".gdf") != 0:
		  raise IOError("Failed to load GDF file: "+bname+".gdf")

		if gdf.WriteBND() != 0:
		  raise IOError("Failed to write BND file: "+bname+".bnd")
		if gdf.WriteBH() != 0:
		  raise IOError("Failed to write BH file: "+bname+".bh")
		if gdf.WriteFEA() != 0:
		  raise IOError("Failed to write FEA file: "+bname+".fea")

		del gdf

		fea = win32com.client.dynamic.Dispatch("pc_fea.AppAutomation")
		if fea.LoadFromFile(bname+".fea") != 0:
		  raise IOError("Failed to load FEA file: "+bname+".fea")

		fea.Pause = 0

		if fea.RunScript() != 0:
		  raise IOError("Failed to run FEA file")

		del fea

	def DoIPSIGoFER(self, arg2, dynamic=None, CSModel=None, ElementTable=None, Transient=None, Cycles=None, CalcTorque=None, Waveform=None, IncShaft=False, FullData=None):
		if self.speed_ext == ".bd4":
			ipsi = self.des.CreateGoFERIPSI

			if isinstance(arg2, str):
				ipsi.UserIWfms = True
				ipsi.UserCurrentFileName = arg2
			else:
				if dynamic != True:
					if self.getParam("ISPSpec") == 0:
						ipsi.MaxISP = self.getParam("ISP")
					elif self.getParam("ISPSpec") == 1:
						ipsi.MaxISP = self.getParam("ISP")*np.sqrt(2.0)
					elif self.getParam("ISPSpec") == 2:
						ipsi.MaxISP = np.sqrt(self.getParam("ISPd")**2 + self.getParam("ISPq")**2)

					if self.getParam("Drive") == 0:
						ipsi.MaxAngle = self.getParam("gamma")
					if self.getParam("Drive") == 1:
						ipsi.MaxAngle = self.getParam("Th0")

					if self.getParam("Drive") == 2: # For Drive=ACVolts we need the o/p values of current and gamma
						ipsi.MaxISP = self.getParam("ILpk")
						if self.getParam("Sw_Ctl") == 9: # Generator
							ipsi.MaxAngle = 180-self.getParam("gammaACV")
						else:
							ipsi.MaxAngle = self.getParam("gammaACV")
					ipsi.ISPCount = 1
					ipsi.AngleCount = 1
					ipsi.DynamicWfms = False
				else:
					ipsi.DynamicWfms = dynamic

				ipsi.RotSteps = arg2
			
			if CSModel != None:
				ipsi.CSModel = CSModel

			if ElementTable != None:
				ipsi.ElementsTable = ElementTable

			if Transient != None:
				if Cycles == None:
					Cycles = 4
				ipsi.Transient = Transient
				ipsi.Cycles = Cycles

			ipsi.IncShaft=IncShaft
			ipsi.NoPostP=True
			ipsi.IndividualMES=True
			ipsi.gdfRunMode=0
			ipsi.Execute()
			self.RunGoFER()

			if CalcTorque != None:
				## read torque from ipsi loop and return torque value
				lines = self.ReadFileLines(self.speed_dir+self.speed_fname+".ltq")
				feaTorque = lines[2].split(' ')
				tindex = self.getParam('IOTorque')
				self.Message("DoIPSIGoFER: i-psi Torque="+str(float(feaTorque[1])*self.torque_factor[tindex]))
				return float(feaTorque[1])*self.torque_factor[tindex]

		elif (self.speed_ext == ".srd") or (self.speed_ext == ".xsrd"):
			ipsi = self.des.CreateGoFERipsi

			if isinstance(arg2, str):
				ipsi.CurrentMode = 2
				ipsi.UserCurrentFileName = arg2
			else:
				ipsi.CurrentMode = 1
				ipsi.RotSteps = arg2
			
			ipsi.gdfRunMode=0
			ipsi.Execute()
			self.RunGoFER()
		else:
			self.Message("DoIPSIGoFER: i-psi GoFER not yet implemented for "+self.appname, err=3)



		if Waveform != None or FullData != None:
			lines = self.ReadFileLines(self.speed_dir+self.speed_fname+"_I1_G1.ipsi")
			RotSteps = len(lines) - 6
			Angle = np.zeros(RotSteps, dtype=float)

			Ipk = float(lines[2].split(' ')[0])
			PhAdv = float(lines[3].split(' ')[0])
			Nph = int(lines[4].split(' ')[0])

			Iph = np.zeros((RotSteps, Nph), dtype=float)
			psi = np.zeros((RotSteps, Nph), dtype=float)
			psiDQ = np.zeros((RotSteps, 2), dtype=float)

			lines = lines[6:]
			i=0
			for line in lines:
				row = line.split(' ')
				Angle[i] = float(row[0])
				for j in range(Nph):
					Iph[i,j] = float(row[j+1])
					psi[i,j] = float(row[j+1+Nph])
				if Nph == 3:
					psiDQ[i,0] = float(row[7])
					psiDQ[i,1] = float(row[8])
				i=i+1

			if Waveform != None:
				return [Angle, Iph[:,0], psi[:,0]]
			else:
				return {'Angle': Angle, 'Iph': Iph, 'psi': psi, 'psiDQ': psiDQ, 'Nphase': Nph, 'Ipeak': Ipk, 'PhaseAdvance': PhAdv}


	def DoCoggingTorqueGoFER(self, RotSteps, plot=None, savefig=None, appendTitle=None, IncShaft=False):
		if self.speed_ext == ".bd4":
			cogg = self.des.CreateGoFERCoggingTorque
			cogg.RotSteps = RotSteps
			cogg.RemoveDC = True
			cogg.Coenergy = True
			cogg.Maxwell = True
			cogg.VW = True
			cogg.MSI = True

			cogg.IncShaft = IncShaft
			cogg.NoPostP=True
			cogg.gdfRunMode=0
			cogg.Execute()
			self.RunGoFER()

		else:
			self.Message("DoCoggingTorqueGoFER: Cogging Torque not yet implemented for "+self.appname, err=3)
			return None

		lines = self.ReadFileLines(self.speed_dir+self.speed_fname+".cog")
		x = np.arange(RotSteps, dtype=float)
		maxwellTorque = x*0.
		coenergyTorque = x*0.
		virtworkTorque = x*0.
		msiTorque = x*0.

		tindex = self.getParam('IOTorque')
		lines = lines[2:]
		i=0
		for line in lines:
			row = line.split(' ')
			x[i] = float(row[0])
			maxwellTorque[i] = float(row[1])*self.torque_factor[tindex]
			coenergyTorque[i] = float(row[4])*self.torque_factor[tindex]
			virtworkTorque[i] = float(row[5])*self.torque_factor[tindex]
			msiTorque[i] = float(row[6])*self.torque_factor[tindex]
			i=i+1

		cog = {'Angle':x, 'Maxwell': maxwellTorque, 'Coenergy': coenergyTorque, 'VW': virtworkTorque, 'MSI': msiTorque, 'trqLabel': self.torque_label[tindex]}

		if appendTitle == None:
			appendTitle = ''
		else:
			appendTitle = ' ('+appendTitle+')'

		if plot != None or savefig !=None:
			plt.figure(1)
			plt.clf()
			plt.plot(cog['Angle'], cog['Coenergy'], 'r')
			plt.plot(cog['Angle'], cog['VW'], 'b')
			plt.plot(cog['Angle'], cog['Maxwell'], 'g')
			plt.xlabel('Theta [deg]')
			plt.ylabel('Cogging Torque ['+cog['trqLabel']+']')
			plt.title('Cogging Torque'+appendTitle)
			plt.legend(['Coenergy', 'Virtual Work', 'Maxwell'])
			plt.xlim([0, np.amax(cog['Angle'])])
			plt.ylim([-2, 2])
		if savefig != None:
			if isinstance(savefig, str):
				fname = savefig
			else:
				fname = self.GetCWD()+"\\" + self.speed_bname+'_CoggTrq.png'
			plt.savefig(fname, bbox_inches='tight')
		if plot !=None:
			plt.show()

		return cog


	def DoBtoothGoFER(self, RotSteps, Waveform=None, BrMode=4, IncShaft=False):
		if self.speed_ext == ".bd4":

			# Do BDC Btooth GoFER
			btooth = self.des.CreateGoFERBtooth
			if btooth.BtRadBtooth == 0:
				self.Static()

			btooth.RotSteps = RotSteps

			btooth.BrMode=BrMode
			btooth.IncShaft=IncShaft

			btooth.NoPostP=True
			btooth.gdfRunMode=0
			ret = btooth.Execute()
			self.RunGoFER()

			destFileName = self.speed_dir + self.speed_fname + ".tfw"

		else:
			self.Message("DoBtoothGoFER: Btooth not yet implemented for "+self.appname, err=3)


		if Waveform != None:
			lines = self.ReadFileLines(destFileName)
			x = np.arange(RotSteps, dtype=float)
			y = x*0.
			lines = lines[4:]
			i=0
			for line in lines:
				row = line.split(' ')
				x[i] = float(row[0])
				y[i] = float(row[1])
				i=i+1

			return [x, y]

	def MatchBtoothIteration(self, XVals, x, y, XNames):
		INPs = {}
		for i in range(len(XVals)):
			INPs[XNames[i]] = XVals[i]

		self.setParams(INPs)
		self.Static()

		XTh = self.getWaveForm('XTh')
		Btooth = self.getWaveForm('BTwfm')

		f = interp.interp1d(XTh, Btooth, kind='cubic')
		devs = y - f(x)

		self.Message("Match Btooth: "+str([np.sum(devs**2)/np.max(np.abs(y)), XVals]))

		return devs/np.max(np.abs(y))

	def MatchBtoothWfm(self, x, y, XAdjust=['XTTarc', 'XBtpk', 'XBetaM'], plot=None, savefig=None, appendTitle=None):

		XTh_0 = self.getWaveForm('XTh')
		Btooth_0 = self.getWaveForm('BTwfm')

		p0 = [self.getParam(param) for param in XAdjust]
		INPs = dict(zip(XAdjust, p0))

		p, pconv = sciopt.leastsq(self.MatchBtoothIteration, p0, args=(x, y, XAdjust))
		OUTs = dict(zip(XAdjust, p))
		self.setParams(OUTs)
		self.Static()

		XTh = self.getWaveForm('XTh')
		Btooth = self.getWaveForm('BTwfm')

		if appendTitle == None:
			appendTitle = ''
		else:
			appendTitle = ' ('+appendTitle+')'

		if plot != None or savefig !=None:
			plt.figure(1)
			plt.clf()
			plt.plot(XTh_0, Btooth_0, 'b')
			plt.plot(x,y, 'ro')
			plt.plot(XTh, Btooth, 'g')
			plt.xlim([0,180])
			plt.xlabel('Angle [edeg]')
			plt.ylabel('Btooth Flux [T]')
			plt.title('ToothFlux Match FE'+appendTitle)
			plt.legend(['Initial Btooth', 'PC-FEA Btooth', 'Matched Btooth'], loc='lower right')
		if savefig != None:
			if isinstance(savefig, str):
				fname = savefig
			else:
				fname = self.GetCWD()+"\\" + self.speed_bname+'_BtoothMatch.png'
			plt.savefig(fname, bbox_inches='tight')
		if plot !=None:
			plt.show()

		self.setParams(INPs)
		return OUTs


	def MatchipsiIteration(self, XVals, x, iph, psi, XNames):
		INPs = {}
		for i in range(len(XVals)):
			INPs[XNames[i]] = XVals[i]

		self.setParams(INPs)
		self.Static()

		XTh = np.array(self.getWaveForm('XTh'))
		Iph1 = np.array(self.getWaveForm('Iw1'))
		Psi1 = np.array(self.getWaveForm('psiw1'))

		f = interp.interp1d(XTh, Iph1, kind='cubic')
		g = interp.interp1d(XTh, Psi1, kind='cubic')

		devs = psi - g(x % 360)*1000

		self.Message("Match i-psi: "+str([np.sum(devs**2)/np.max(np.abs(psi)), XVals]))
		return np.sum(devs**2)/np.max(np.abs(psi))


	def MatchipsiWfm(self, x, y, z, XAdjust=['XCd', 'XCq', 'XBrT'], bounds=None, plot=None, savefig=None, appendTitle=None, method='SLSQP', tol=0.01, eps=0.01):

		XTh_0 = np.array(self.getWaveForm('XTh'))
		Iph1_0 = np.array(self.getWaveForm('Iw1'))
		Psi1_0 = np.array(self.getWaveForm('psiw1'))

		p0 = [self.getParam(param) for param in XAdjust]
		INPs = dict(zip(XAdjust, p0))

		opts = {"maxiter": 4, "eps": eps, "ftol": tol}
		if bounds == None:
			bounds = [[0.1, 3] for name in XAdjust]
			if 'XBrT' in XAdjust:
				bounds[XAdjust.index('XBrT')] = [0.8, 1.2]


		p = sciopt.minimize(self.MatchipsiIteration, p0, args=(x, y, z, XAdjust), options=opts, method=method, bounds=bounds)
		OUTs = dict(zip(XAdjust, p.x))
		self.setParams(OUTs)
		self.Static()

		XTh = np.array(self.getWaveForm('XTh'))
		Iph1 = np.array(self.getWaveForm('Iw1'))
		Psi1 = np.array(self.getWaveForm('psiw1'))

		if appendTitle == None:
			appendTitle = ''
		else:
			appendTitle = ' ('+appendTitle+')'

		if plot != None or savefig !=None:
			plt.figure(1)
			plt.clf()
			plt.plot(Iph1_0, Psi1_0*1000, 'b')
			plt.plot(y, z, 'ro')
			plt.plot(Iph1, Psi1*1000, 'g')
			plt.xlabel('Iph1 [A]')
			plt.ylabel('Flux Linkage Ph1 [mV-s]')
			plt.title('i-psi Loop'+appendTitle)
			plt.legend(['Initial i-psi', 'PC-FEA i-psi', 'Matched i-psi'], loc='lower left')
		if savefig != None:
			if isinstance(savefig, str):
				fname = savefig
			else:
				fname = self.GetCWD()+"\\" + self.speed_bname+'_IPsiMatch.png'
			plt.savefig(fname, bbox_inches='tight')
		if plot !=None:
			plt.show()

		self.setParams(INPs)
		return OUTs

	def AutoSearchFunc(self, x, pVar, pMatch, matchVal):
		INPs = {}

		if isinstance(pVar, list):
			for i in range(len(x)):
				INPs[pVar[i]] = x[i]
		else:
			INPs[pVar] = x
		self.setParams(INPs)
		self.Dynamic()
		val = self.getParam(pMatch)
		if isinstance(pVar, list):
			self.Message(" ".join(pVar)+": "+str(x)+pMatch+": "+str(val))
		else:
			self.Message(pVar+": "+str(x)+pMatch+": "+str(val))
		
		if isinstance(matchVal, str):
			if matchVal.lower().startswith("max"):
				return (-val)
			else:
				return (val)
		else:
			return (matchVal-val)**2

	def MaxEffMapCost(self, x, pVar, pMatch, MatchVal):
		INPs = {}

		if isinstance(pVar, list):
			for i in range(len(x)):
				INPs[pVar[i]] = x[i]
		else:
			INPs[pVar] = x
		self.setParams(INPs)
		self.Static()
		val = self.getParam(pMatch)
		VLL1 = self.getParam("VLL1")
		Vct1 = self.getParam("Vct1")
		Eff = self.getParam("Eff")

		if isinstance(MatchVal, str):
			if MatchVal.lower().startswith("max"):
				cost = -10*val
			else:
				cost = 10*val
		else:
			cost = 10*(val-MatchVal)**2

		cost = cost + (max(VLL1-1.03*Vct1, 0))**2 - Eff

		if isinstance(pVar, list):
			self.Message(" ".join(pVar)+": "+str(x)+", "+pMatch+": "+str(val)+" ["+str(MatchVal)+"], VLL1: "+str(VLL1)+" ["+str(1.03*Vct1)+"], Eff: "+str(Eff))
		else:
			self.Message(pVar+": "+str(x)+", "+pMatch+": "+str(val)+" ["+str(MatchVal)+"], VLL1: "+str(VLL1)+" ["+str(1.03*Vct1)+"], Eff: "+str(Eff))

		return cost

	def DoAutoSearch(self, pVar, pMatch, MatchVal, Tol=0.01, bounds=None, userFunc=None, method='SLSQP', maxIter=50):
		argTuple = (pVar, pMatch, MatchVal)

		if userFunc == None:
			userFunc = self.AutoSearchFunc

		if isinstance(pVar, str):
			# this is univariate minimization, use minimize_scalar
			if bounds == None:
				res = sciopt.minimize_scalar(userFunc, args=argTuple, tol=Tol)
			else:
				if isinstance(MatchVal, str):
					opts = {'xatol': Tol, 'maxiter': maxIter}
				else:
					opts = {'xatol': MatchVal*Tol**2, 'maxiter': maxIter}
				res = sciopt.minimize_scalar(userFunc, args=argTuple, options=opts, method='bounded', bounds=bounds)
		else:
			# this is multivariate minimization, use normal minimize function in sciopt

			if len(pVar) == 1:
				pVar = pVar[0]
				varZero = self.getParam(pVar)
			else:
				varZero = np.zeros(len(pVar))
				for i in range(len(varZero)):
					varZero[i] = self.getParam(pVar[i])

			opts = {"maxiter": maxIter, "eps": 1}

			if bounds == None:
				res = sciopt.minimize(userFunc, varZero, args=argTuple, method=method, options=opts)
			else:
				res = sciopt.minimize(userFunc, varZero, args=argTuple, method=method, bounds=bounds, options=opts)

		return res

	def ValidateGeometry(self):

		if self.speed_ext == '.bd4':
			valGeom = self.des.Validate()
			if valGeom < 0:
				self.Message("ValidateGeometry: " + self.des.GetValidateErrorMsg(valGeom) + "("+str(valGeom)+")", err=-1)
			return valGeom
		else:
			self.Message("This SPEED module ("+self.appname+") does not support ValidateGeometry()", err=1)
			return 101


	def WriteLamBMP(self, fileName=None, options=None):
		if fileName == None:
			fileName = self.GetCWD() + "\\" + self.speed_fname + "_updatedLamDesign.bmp"
		if options == None:
			options = "size=200, background=white, rotorfill=1, statorfill=1, magnetfill=true, rotor=lightsteelblue, stator=silver, north=red, south=green "
	
		self.des.WriteThumbnail(fileName, options )

		
	def Map2DCalc(self, rpm, load, hist=None, current=None, Tol=0.01, mode=0, ipsiCalc=1, nCycles=3):
		# rpm and load should be positive definite (use mode to specify motoring or generating or both
		# mode=0 - motoring only
		# mode=1 - generating only
		# mode=2 - both motoring and generating side of the map

		# Make rpm and load have the proper 1D form in case they are in meshgrid form
		if len(rpm.shape) == 2:
			rpm = rpm[0,:]
		if len(load.shape) == 2:
			load = load[:,0]

		# create tuple for dimensions of all 2D maps
		if mode==2:
			N_load = 2*len(load)
		else:
			N_load = len(load)
		dim = (N_load, len(rpm))
		self.Message("Map2DCalc: Dim="+str(dim))

		if hist != None:
			if hist.shape != dim:
				self.Message("Map2DCalc: Invalid use of hist, dimensions do not agree with rpm/load", err=-1)
		else:
			hist = np.ones(dim)

		# Define the map structures
		map = Map2D(dim)
		self.Message("Performing Map2DCalc: Dim="+str(dim))

		if self.getParam("gamma") < 90:
			gammaMot = self.getParam("gamma")
		else:
			gammaMot = 180 - self.getParam("gamma")

		for j in range(0, len(load), 1):
			if current:
				self.setParams({"ISP": load[j]})
			for i in range(0, len(rpm), 1):
				self.setParams({"RPM": rpm[i]})
				if mode == 2:
					J = len(load)+j
					N = len(load)-1-j
				else:
					J = j
					N = j

				# Motoring Cases
				if (mode == 0 or mode == 2) and hist[J,i] > 0:
					it_time = time.time()
					self.setParams({"gamma": gammaMot})
					self.Message("RPM="+str(rpm[i])+"\tLoad="+str(load[j])+"\tnHist="+str(hist[J,i]))

					self.setParams({"ipsiCalc": ipsiCalc})
					sol = self.SearchTorque(load[j], current=current, Tol=Tol, nCycles=nCycles)
					if sol.success:
					  map.SetFromSPEED(J, i, self, target=load[[j]])
					else:
					  map.SetFromSPEED0(J, i, self, target=load[[j]])

					self.Message("Result: "+str([map.rpm[J,i], map.load_target[J,i], map.TShaft[J,i], map.isp[J,i], map.gamma[J,i], map.PShaft[J,i], map.Eff[J,i]]), err=0)
					gammaMot = self.getParam("gamma")
					self.Message("Iteration Execution Time="+str(time.time() - it_time), err=0)

				else:
					if hist[J,i] == 0:
						self.Message("Skipping Load="+str(load[j])+", 0 count")



				# Generating Cases
				if (mode == 1 or mode == 2) and hist[N,i] > 0:
					it_time = time.time()
					if self.getParam("gamma") < 90:
						self.setParams({"gamma": 180-self.getParam("gamma")})
					self.Message("RPM="+str(rpm[i])+"\tLoad="+str(-load[j])+"\tnHist="+str(hist[N,i]))

					self.setParams({"ipsiCalc": ipsiCalc})
					self.SearchTorque(-load[j], current=current, Tol=Tol, nCycles=nCycles, Gen=True)

					map.SetFromSPEED(N, i, self, target=-load[[j]])

					self.Message("Result: "+str([map.rpm[N,i], map.load_target[N,i], map.TShaft[N,i], map.isp[N,i], map.gamma[N,i], map.PShaft[N,i], map.Eff[N,i]]), err=0)
					self.Message("Iteration Execution Time="+str(time.time() - it_time), err=0)
				else:
					if hist[N,i] == 0:
						self.Message("Skipping Trq="+str(load[j])+", 0 count")

		return map

	
	def MatchTorque(self, TrqTarget, Tol=0.01, URF=0.8, maxiter=10):
		self.Dynamic()
		Trq = self.getParam('Tshaft')
		isp = self.getParam('ISP')
		it_num=0
		while(abs((Trq-TrqTarget)/TrqTarget) > Tol and it_num < maxiter):
			isp = self.getParam('ISP')*((TrqTarget-Trq)*URF + Trq)/Trq
			self.setParams({'ISP': isp})
			self.Dynamic()
			Trq = self.getParam('Tshaft')
			self.Message("MatchTorque: ISP="+str(isp) + "\tTrq="+str(Trq))
			it_num = it_num + 1

		return isp


	def SearchTorque(self, target, nCycles=3, current=None, Tol=0.1, Gen=False):
		ipsiCalc = self.getParam("ipsiCalc")
		if ipsiCalc == 0:
			nCycles = 1

		for i in range(nCycles):
			self.setParams({"ipsiCalc": 0})

			if current:
				if Gen:
					bounds = [92, 180]
				else:
					bounds = [0, 88]
				sol = self.DoAutoSearch("gamma", "Tshaft", "max", userFunc=self.MaxEffMapCost, bounds=bounds, maxIter=7-2*i)
				self.setParams({"gamma": sol.x})
			else:
				if Gen:
					bounds = [[0, 450], [92, 180]]
				else:
					bounds = [[0, 450], [0, 88]]
				sol = self.DoAutoSearch(["ISP", "gamma"], "Tshaft", target, userFunc=self.MaxEffMapCost, bounds=bounds, maxIter=7-2*i)

			if (ipsiCalc != 0 and i != nCycles-1):
				self.setParams({"ipsiCalc": ipsiCalc})
				res = self.Static()
				if res < 0:
					self.Message("SearchTorque(): Static design failed - "+self.des.GetErrorMsg(res))

			return sol

			



	def HEEDSOutput(self, varName, val=None):
		if self.heeds_out != None:
			closed = self.heeds_out.closed
		else:
			closed = True

		if closed:
			try:
				self.heeds_out = open(self.GetCWD()+"\\heeds.output", "w")
			except (IOError, WindowsError):
				self.Message("Cannot Open heeds.output", err=-1)

		if val == None:
			val = self.getParam(varName)
			varName = self.getParamName(varName)
		elif isinstance(val, str):
			val = self.getParam(val)

		self.heeds_out.write(varName + (" = %10.10E\n" % (val)))
		return val

	def GetCWD(self):
		return os.getcwd()

	def Save(self, appendName=None):
		if appendName == None:
			appendName = ""
		self.des.SaveToFile(self.speed_dir + self.speed_fname + appendName + self.speed_ext)

	def SaveCWD(self, appendName=None):
		if appendName == None:
			appendName = ""
		self.des.SaveToFile(self.GetCWD()+"\\" + self.speed_fname + appendName + self.speed_ext)

	def SaveAs(self, fileout):
		self.des.SaveToFile(fileout)

	def ReadFileLines(self, filename):
		self.Message("ReadFileLines: Attempting to read file "+filename)
		return [re.sub(' +',' ',line.strip()) for line in open(filename, 'r')]

	def __del__(self):
		if hasattr(self, 'wapp'):
			self.wapp.Quit 
			del self.wapp
		if hasattr(self, 'wpi'):
			del self.wpi
		if hasattr(self, 'des'):
			del self.des
		CheckClean()

		if not self.log_file.closed:
			self.log_file.close()

		if self.heeds_out != None:
			if not self.heeds_out.closed:
				self.heeds_out.close()

		## for HEEDS compatibility, write file heeds.complete when speedWrapper exits
		try:
			fout = open(self.GetCWD()+"\\heeds.complete", "wt", encoding='utf-8')
			fout.write("speedWrapper has exited")
			fout.close()
		except (IOError, WindowsError):
			self.Message("Cannot Open heeds.output", err=-1)


	def updatePIparams(self, wapp, fname):
		filepath = shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0)
		if filepath == None:
			filepath = os.getenv('LOCALAPPDATA', None)
		filepath = filepath + '\\SPEED'
		
		# make sure the PI directory is in the Python path
		if not filepath in sys.path:
			sys.path.append(filepath)
			
		updateNeeded = 0
		try:
			m = self.parsePIparams(filepath+"\\"+fname+".py")
		except IOError:
			updateNeeded = 1
			
		if not updateNeeded:
			updateNeeded = (m['VERSIONNUMBER'] != wapp.VersionNumber)
	
		if updateNeeded:
			res = wapp.WriteParameterInformation(filepath+"\\"+fname+".py", "*.py")
			m = self.parsePIparams(filepath+"\\"+fname+".py")

		return m


	def parsePIparams(self, fullpathfilename):
		piParams = {}
		with open(fullpathfilename, "r") as f:
			for line in f:
				key,val = line.split("=")
				if val.strip()[0] == '"':
					piParams[key.strip()] = val.strip(' "\n')
				else:
					piParams[key.strip()] = int(val.strip(), 16)
		return piParams
			
	def strParseNum(self, inStr):
		try:
			return float(inStr)
		except Exception:
			return inStr



class Map2D:

	' Object to hold 2D Maps of performance data'

	def __init__(self, dim):
		self.load_target = np.zeros(dim, dtype=float)
		self.rpm = np.zeros(dim, dtype=float)
		self.isp = np.zeros(dim, dtype=float)
		self.gamma = np.zeros(dim, dtype=float)
		self.XCq = np.zeros(dim, dtype=float)
		self.XCd = np.zeros(dim, dtype=float)
		self.XBrT = np.zeros(dim, dtype=float)

		self.TShaft = np.zeros(dim, dtype=float)
		self.PShaft = np.zeros(dim, dtype=float)
		self.PElec = np.zeros(dim, dtype=float)
		self.Eff = np.zeros(dim, dtype=float)
		self.WCu = np.zeros(dim, dtype=float)
		self.WFe = np.zeros(dim, dtype=float)

	def SetFromSPEED(self, i, j, spInst, target=None):
		if target != None:
			self.load_target[i,j] = target
		self.rpm[i,j] = spInst.getParam("RPM")
		self.isp[i,j] = spInst.getParam("ISP")
		self.gamma[i,j] = spInst.getParam("gamma")
		self.XCq[i,j] = spInst.getParam("XCq")
		self.XCd[i,j] = spInst.getParam("XCd")
		self.XBrT[i,j] = spInst.getParam("XBrT")

		self.TShaft[i,j] = spInst.getParam("TShaft")
		self.PShaft[i,j] = spInst.getParam("PShaft")
		self.PElec[i,j] = spInst.getParam("PElec")
		self.Eff[i,j] = spInst.getParam("EFF")
		self.WCu[i,j] = spInst.getParam("WCu")
		self.WFe[i,j] = spInst.getParam("WIron")

	def SetFromSPEED0(self, i, j, spInst, target=None):
		if target != None:
			self.load_target[i,j] = target
		self.rpm[i,j] = 0
		self.isp[i,j] = 0
		self.gamma[i,j] = 0
		self.XCq[i,j] = 0
		self.XCd[i,j] = 0
		self.XBrT[i,j] = 0

		self.TShaft[i,j] = 0
		self.PShaft[i,j] = 0
		self.PElec[i,j] = 0
		self.Eff[i,j] = 0
		self.WCu[i,j] = 0
		self.WFe[i,j] = 0

	def EffMap(self, spInst, plot=None, savefig=None, appendTitle=None, interpolation='nearest', clim=None, extent=None, contour=False):
		if appendTitle == None:
			appendTitle = ''
		else:
			appendTitle = ' ('+appendTitle+')'

		if plot != None or savefig !=None:
			fig = plt.figure()
			ax = fig.add_subplot(111)
			if extent == None:
				extent = (np.min(self.rpm), np.max(self.rpm), np.min(self.TShaft), np.max(self.TShaft))
			if contour:
				if clim != None:
					v = np.linspace(clim[0], clim[1], 64, endpoint=True)
					im = plt.contourf(self.rpm, self.TShaft, self.Eff, 64, cmap=cm.jet)
					im.set_clim(clim[0], clim[1])
					plt.colorbar(im, ticks=v)
				else:
					im = plt.contourf(self.rpm, self.TShaft, self.Eff, 64, cmap=cm.jet)
					plt.colorbar(im)
			else:
				im = ax.imshow(self.Eff, interpolation=interpolation, cmap=cm.jet, extent=extent)
				if clim != None:
					im.set_clim(clim[0], clim[1])
				plt.colorbar(im)

			ax.set_aspect('auto')
			plt.xlabel('RPM')
			plt.ylabel('Torque ['+spInst.torque_label[spInst.getParam('IOTorque')]+']')
			plt.title('Efficiency Map '+appendTitle)

		if savefig != None:
			if isinstance(savefig, str):
				fname = savefig
			else:
				fname = spInst.GetCWD()+"\\" + spInst.speed_bname +'_EffMap.png'
			plt.savefig(fname, bbox_inches='tight')
		if plot !=None:
			plt.show()
