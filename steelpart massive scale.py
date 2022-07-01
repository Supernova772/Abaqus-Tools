#!/user/bin/python
#-*-coding:UTF-8-*-
from abaqus import *
from abaqusConstants import *
from caeModules import *
from driverUtils import executeOnCaeStartup
from part import *
from material import *
from section import *
from assembly import *
import regionToolset
from step import *
from interaction import *
from load import *
from mesh import *
from optimization import *
from job import *
from sketch import *
from visualization import *
from connectorBehavior import *
import matplotlib.pyplot as plt
import numpy as np
from odbAccess import *
import xlwt
import time
import math
import xlwings as xw
import threading
for h in range(5): 
    for o in range(2):
        for l in range(5):
            for n in range (5):
                for m in range(2):
                    dep=300+50*h
                    wid=100+10*o
                    ft=10+2*l
                    wt=5+n
                    it=5+m
                    dep=str(dep)
                    wid=str(wid)
                    ft=str(ft)
                    wt=str(wt)
                    it=str(it)
                    mdbname='ub'+dep+'x'+wid+'x'+ft+'x'+wt+'it'+it
                    dep=int(dep)
                    wid=int(wid)
                    ft=int(ft)
                    wt=int(wt)
                    ic=int(ic)
                    insu=int(it)
                    mdb.Model(name=mdbname, absoluteZero=-273.15, stefanBoltzmann=5.67e-14, 
                        modelType=STANDARD_EXPLICIT)
                    s = mdb.models[mdbname].ConstrainedSketch(name='__profile__', 
                        sheetSize=200.0)
                    g, v, d, c = s.geometry, s.vertices, s.dimensions, s.constraints
                    s.setPrimaryObject(option=STANDALONE)
                    session.viewports['Viewport: 1'].view.setValues(nearPlane=163.741, 
                        farPlane=213.383, width=249.945, height=115.305, cameraPosition=(8.20491, 
                        13.233, 188.562), cameraTarget=(8.20491, 13.233, 0))
                    s.Line(point1=(-1*wid/2, dep/2), point2=(wid/2, dep/2))
                    s.HorizontalConstraint(entity=g[2], addUndoState=False)
                    s.Line(point1=(wid/2, dep/2), point2=(wid/2, dep/2-ft))
                    s.VerticalConstraint(entity=g[3], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[2], entity2=g[3], addUndoState=False)
                    s.Line(point1=(wid/2, dep/2-ft), point2=(wt/2, dep/2-ft))
                    s.HorizontalConstraint(entity=g[4], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[3], entity2=g[4], addUndoState=False)
                    s.Line(point1=(wt/2, dep/2-ft), point2=(wt/2, -1*(dep/2-ft)))
                    s.VerticalConstraint(entity=g[5], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[4], entity2=g[5], addUndoState=False)
                    s.Line(point1=(wt/2, -1*(dep/2-ft)), point2=(wid/2, -1*(dep/2-ft)))
                    s.HorizontalConstraint(entity=g[6], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[5], entity2=g[6], addUndoState=False)
                    s.Line(point1=(wid/2, -1*(dep/2-ft)), point2=(wid/2, -1*dep/2))
                    s.VerticalConstraint(entity=g[7], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[6], entity2=g[7], addUndoState=False)
                    s.Line(point1=(wid/2, -1*dep/2), point2=(-1*wid/2, -1*dep/2))
                    s.HorizontalConstraint(entity=g[8], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[7], entity2=g[8], addUndoState=False)
                    s.Line(point1=(-1*wid/2, -1*dep/2), point2=(-1*wid/2, -1*(dep/2-ft)))
                    s.VerticalConstraint(entity=g[9], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[8], entity2=g[9], addUndoState=False)
                    s.Line(point1=(-1*wid/2, -1*(dep/2-ft)), point2=(-1*wt/2, -1*(dep/2-ft)))
                    s.HorizontalConstraint(entity=g[10], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[9], entity2=g[10], addUndoState=False)
                    s.Line(point1=(-1*wt/2, -1*(dep/2-ft)), point2=(-1*wt/2, dep/2-ft))
                    s.VerticalConstraint(entity=g[11], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[10], entity2=g[11], addUndoState=False)
                    s.Line(point1=(-1*wt/2, dep/2-ft), point2=(-1*wid/2, dep/2-ft))
                    s.HorizontalConstraint(entity=g[12], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[11], entity2=g[12], addUndoState=False)
                    s.Line(point1=(-1*wid/2, dep/2-ft), point2=(-1*wid/2, dep/2))
                    s.VerticalConstraint(entity=g[13], addUndoState=False)
                    s.PerpendicularConstraint(entity1=g[12], entity2=g[13], addUndoState=False)
                    p = mdb.models[mdbname].Part(name='steel', dimensionality=TWO_D_PLANAR, 
                        type=DEFORMABLE_BODY)
                    p = mdb.models[mdbname].parts['steel']
                    p.BaseShell(sketch=s)
                    mdb.models[mdbname].parts['steel'].BaseShell(sketch=
                        mdb.models[mdbname].sketches['__profile__'])
                    s.unsetPrimaryObject()
                    p = mdb.models[mdbname].parts['steel']
                    session.viewports['Viewport: 1'].setValues(displayedObject=p)
                    del mdb.models[mdbname].sketches['__profile__']

                    #dep,wid,ft,wt,insu= getInputs(fields=(('Depth:','351.4'),('Width:','171.1'),('Flange thickness:','9.7'),
                    #('Web thickness:','7'),('Insulation thickness:',8)),label='Enter parameters',dialogTitle='Steel section parameters')

                    #Fire resistance material
                    s1 = mdb.models[mdbname].ConstrainedSketch(name='__profile__', 
                        sheetSize=200.0)
                    g, v, d, c = s1.geometry, s1.vertices, s1.dimensions, s1.constraints
                    s1.Line(point1=(-1*wid/2, dep/2), point2=(-1*(wid/2+insu), dep/2))
                    s1.HorizontalConstraint(entity=g[2], addUndoState=False)
                    session.viewports['Viewport: 1'].view.setValues(nearPlane=45.8917, 
                        farPlane=331.232, width=1625.93, height=750.075, cameraPosition=(379.37, 
                        -97.4115, 188.562), cameraTarget=(379.37, -97.4115, 0))
                    s1.Line(point1=(-1*(wid/2+insu), dep/2), point2=(-1*(wid/2+insu), dep/2-ft-insu))
                    s1.VerticalConstraint(entity=g[3], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[2], entity2=g[3], addUndoState=False)
                    s1.Line(point1=(-1*(wid/2+insu), dep/2-ft-insu), point2=(-1*(wt/2+insu), dep/2-ft-insu))
                    s1.HorizontalConstraint(entity=g[4], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[3], entity2=g[4], addUndoState=False)
                    s1.Line(point1=(-1*(wt/2+insu), dep/2-ft-insu), point2=(-1*(wt/2+insu), -1*(dep/2-ft-insu)))
                    s1.VerticalConstraint(entity=g[5], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[4], entity2=g[5], addUndoState=False)
                    s1.Line(point1=(-1*(wt/2+insu), -1*(dep/2-ft-insu)), point2=(-1*(wid/2+insu), -1*(dep/2-ft-insu)))
                    s1.HorizontalConstraint(entity=g[6], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[5], entity2=g[6], addUndoState=False)
                    s1.Line(point1=(-1*(wid/2+insu), -1*(dep/2-ft-insu)), point2=(-1*(wid/2+insu), -1*(dep/2+insu)))
                    s1.VerticalConstraint(entity=g[7], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[6], entity2=g[7], addUndoState=False)
                    s1.Line(point1=(-1*(wid/2+insu), -1*(dep/2+insu)), point2=(wid/2+insu, -1*(dep/2+insu)))
                    s1.HorizontalConstraint(entity=g[8], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[7], entity2=g[8], addUndoState=False)
                    s1.Line(point1=(wid/2+insu, -1*(dep/2+insu)), point2=(wid/2+insu, -1*(dep/2-ft-insu)))
                    s1.VerticalConstraint(entity=g[9], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[8], entity2=g[9], addUndoState=False)
                    s1.Line(point1=(wid/2+insu, -1*(dep/2-ft-insu)), point2=(wt/2+insu, -1*(dep/2-ft-insu)))
                    s1.HorizontalConstraint(entity=g[10], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[9], entity2=g[10], addUndoState=False)
                    s1.Line(point1=(wt/2+insu, -1*(dep/2-ft-insu)), point2=(wt/2+insu, dep/2-ft-insu))
                    s1.VerticalConstraint(entity=g[11], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[10], entity2=g[11], addUndoState=False)
                    s1.Line(point1=(wt/2+insu, dep/2-ft-insu), point2=(wid/2+insu, dep/2-ft-insu))
                    s1.HorizontalConstraint(entity=g[12], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[11], entity2=g[12], addUndoState=False)
                    s1.Line(point1=(wid/2+insu, dep/2-ft-insu), point2=(wid/2+insu, dep/2))
                    s1.VerticalConstraint(entity=g[13], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[12], entity2=g[13], addUndoState=False)
                    s1.Line(point1=(wid/2+insu, dep/2), point2=(wid/2, dep/2))
                    s1.HorizontalConstraint(entity=g[14], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[13], entity2=g[14], addUndoState=False)
                    s1.Line(point1=(wid/2,dep/2), point2=(wid/2, dep/2-ft))
                    s1.VerticalConstraint(entity=g[15], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[14], entity2=g[15], addUndoState=False)
                    s1.Line(point1=(wid/2, dep/2-ft), point2=(wt/2, dep/2-ft))
                    s1.HorizontalConstraint(entity=g[16], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[15], entity2=g[16], addUndoState=False)
                    s1.Line(point1=(wt/2, dep/2-ft), point2=(wt/2, -1*(dep/2-ft)))
                    s1.VerticalConstraint(entity=g[17], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[16], entity2=g[17], addUndoState=False)
                    s1.Line(point1=(wt/2, -1*(dep/2-ft)), point2=(wid/2, -1*(dep/2-ft)))
                    s1.HorizontalConstraint(entity=g[18], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[17], entity2=g[18], addUndoState=False)
                    s1.Line(point1=(wid/2, -1*(dep/2-ft)), point2=(wid/2, -1*dep/2))
                    s1.VerticalConstraint(entity=g[19], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[18], entity2=g[19], addUndoState=False)
                    s1.Line(point1=(wid/2, -1*dep/2), point2=(-1*(wid/2), -1*dep/2))
                    s1.HorizontalConstraint(entity=g[20], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[19], entity2=g[20], addUndoState=False)
                    s1.Line(point1=(-1*(wid/2), -1*dep/2), point2=(-1*(wid/2), -1*(dep/2-ft)))
                    s1.VerticalConstraint(entity=g[21], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[20], entity2=g[21], addUndoState=False)
                    s1.Line(point1=(-1*(wid/2), -1*(dep/2-ft)), point2=(-1*(wt/2), -1*(dep/2-ft)))
                    s1.HorizontalConstraint(entity=g[22], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[21], entity2=g[22], addUndoState=False)
                    s1.Line(point1=(-1*(wt/2), -1*(dep/2-ft)), point2=(-1*(wt/2), dep/2-ft))
                    s1.VerticalConstraint(entity=g[23], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[22], entity2=g[23], addUndoState=False)
                    s1.Line(point1=(-1*(wt/2),dep/2-ft), point2=(-1*(wid/2), dep/2-ft))
                    s1.HorizontalConstraint(entity=g[24], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[23], entity2=g[24], addUndoState=False)
                    s1.Line(point1=(-1*(wid/2), dep/2-ft), point2=(-1*(wid/2), dep/2))
                    s1.VerticalConstraint(entity=g[25], addUndoState=False)
                    s1.PerpendicularConstraint(entity1=g[24], entity2=g[25], addUndoState=False)
                    session.viewports['Viewport: 1'].view.setValues(nearPlane=480.371, 
                        farPlane=724.523, width=1229.28, height=567.096, cameraPosition=(215.272, 
                        37.1672, 602.447), cameraTarget=(215.272, 37.1672, 0))
                    p1 = mdb.models[mdbname].Part(name='Fire resistance', 
                        dimensionality=TWO_D_PLANAR, type=DEFORMABLE_BODY)
                    p1 = mdb.models[mdbname].parts['Fire resistance']
                    p1.BaseShell(sketch=s1)
                    mdb.models[mdbname].parts['Fire resistance'].BaseShell(sketch=
                        mdb.models[mdbname].sketches['__profile__'])
                    s1.unsetPrimaryObject()
                    p1 = mdb.models[mdbname].parts['Fire resistance']
                    session.viewports['Viewport: 1'].setValues(displayedObject=p1)
                    del mdb.models[mdbname].sketches['__profile__']

                    from material import createMaterialFromDataString
                    createMaterialFromDataString(mdbname, 'Concrete', '2021', 
                        """{'description': '', 'density': {'temperatureDependency': ON, 'table': ((2.34931e-06, 20.0), (2.33992e-06, 60.0), (2.33053e-06, 100.0), (2.32584e-06, 120.0), (2.32114e-06, 140.0), (2.31645e-06, 160.0), (2.31175e-06, 180.0), (2.30706e-06, 200.0), (2.30237e-06, 220.0), (2.29767e-06, 240.0), (2.29298e-06, 260.0), (2.28828e-06, 280.0), (2.28359e-06, 300.0), (2.2789e-06, 320.0), (2.2742e-06, 340.0), (2.26951e-06, 360.0), (2.26481e-06, 380.0), (2.26012e-06, 400.0), (2.25073e-06, 440.0), (2.24134e-06, 480.0), (2.23196e-06, 520.0), (2.22257e-06, 560.0), (2.21318e-06, 600.0), (2.20379e-06, 640.0), (2.1944e-06, 680.0), (2.18502e-06, 720.0), (2.17563e-06, 760.0), (2.16624e-06, 800.0), (2.15685e-06, 840.0), (2.14746e-06, 880.0), (2.13808e-06, 920.0), (2.12869e-06, 960.0), (2.1193e-06, 1000.0), (2.10991e-06, 1040.0), (2.10052e-06, 1080.0), (2.09114e-06, 1120.0), (2.08175e-06, 1160.0), (2.07236e-06, 1200.0)), 'dependencies': 0, 'fieldName': '', 'distributionType': UNIFORM}, 'specificHeat': {'temperatureDependency': ON, 'table': ((900.0, 20.0), (900.0, 60.0), (900.0, 100.0), (2531.4, 115.0), (1000.0, 200.0), (1010.0, 220.0), (1020.0, 240.0), (1030.0, 260.0), (1040.0, 280.0), (1050.0, 300.0), (1060.0, 320.0), (1070.0, 340.0), (1080.0, 360.0), (1090.0, 380.0), (1100.0, 400.0), (1100.0, 440.0), (1100.0, 480.0), (1100.0, 520.0), (1100.0, 560.0), (1100.0, 600.0), (1100.0, 640.0), (1100.0, 680.0), (1100.0, 720.0), (1100.0, 760.0), (1100.0, 800.0), (1100.0, 840.0), (1100.0, 880.0), (1100.0, 920.0), (1100.0, 960.0), (1100.0, 1000.0), (1100.0, 1040.0), (1100.0, 1080.0), (1100.0, 1120.0), (1100.0, 1160.0), (1100.0, 1200.0)), 'dependencies': 0, 'law': CONSTANTVOLUME}, 'materialIdentifier': '', 'conductivity': {'temperatureDependency': ON, 'table': ((0.001642218, 20.0), (0.001568622, 60.0), (0.00149765, 100.0), (0.001429302, 140.0), (0.001363578, 180.0), (0.001300478, 220.0), (0.001240002, 260.0), (0.00118215, 300.0), (0.001126922, 340.0), (0.001074318, 380.0), (0.001024338, 420.0), (0.000976982, 460.0), (0.00093225, 500.0), (0.000890142, 540.0), (0.000850658, 580.0), (0.000813798, 620.0), (0.000779562, 660.0), (0.00074795, 700.0), (0.000718962, 740.0), (0.000692598, 780.0), (0.000668858, 820.0), (0.000647742, 860.0), (0.00062925, 900.0), (0.000613382, 940.0), (0.000600138, 980.0), (0.000589518, 1020.0), (0.000581522, 1060.0), (0.00057615, 1100.0), (0.000573402, 1140.0), (0.000573278, 1180.0), (0.0005742, 1200.0)), 'dependencies': 0, 'type': ISOTROPIC}, 'name': 'Concrete'}""")
                    from material import createMaterialFromDataString
                    createMaterialFromDataString(mdbname, 'Fire resistance material', 
                        '6-14', 
                        """{'description': '', 'density': {'temperatureDependency': OFF, 'table': ((2.4e-07,),), 'dependencies': 0, 'fieldName': '', 'distributionType': UNIFORM}, 'specificHeat': {'temperatureDependency': OFF, 'table': ((1047.0,),), 'dependencies': 0, 'law': CONSTANTVOLUME}, 'materialIdentifier': '', 'conductivity': {'temperatureDependency': OFF, 'table': ((2e-05,),), 'dependencies': 0, 'type': ISOTROPIC}, 'name': 'Fire resistance material'}""")
                    from material import createMaterialFromDataString
                    createMaterialFromDataString(mdbname, 'Steel', '2021', 
                        """{'description': '', 'density': {'temperatureDependency': OFF, 'table': ((7.9e-06,),), 'dependencies': 0, 'fieldName': '', 'distributionType': UNIFORM}, 'specificHeat': {'temperatureDependency': ON, 'table': ((439.8, 20.0), (465.77, 60.0), (487.62, 100.0), (506.18, 140.0), (522.33, 180.0), (536.9, 220.0), (550.75, 260.0), (564.74, 300.0), (579.71, 340.0), (596.51, 380.0), (616.01, 420.0), (639.06, 460.0), (666.5, 500.0), (699.18, 540.0), (737.97, 580.0), (760.21, 600.0), (805.8, 645.0), (936.87, 690.0), (5000.0, 735.0), (1001.92, 770.0), (785.81, 805.0), (708.48, 840.0), (668.75, 875.0), (650.0, 900.0), (650.0, 1200.0)), 'dependencies': 0, 'law': CONSTANTVOLUME}, 'materialIdentifier': '', 'conductivity': {'temperatureDependency': ON, 'table': ((0.05333, 20.0), (0.052, 60.0), (0.05067, 100.0), (0.04933, 140.0), (0.048, 180.0), (0.04667, 220.0), (0.04534, 260.0), (0.04401, 300.0), (0.04267, 340.0), (0.04134, 380.0), (0.04001, 420.0), (0.03868, 460.0), (0.03734, 500.0), (0.03601, 540.0), (0.03468, 580.0), (0.03401, 600.0), (0.03252, 645.0), (0.03102, 690.0), (0.02952, 735.0), (0.02835, 770.0), (0.0273, 805.0), (0.0273, 840.0), (0.0273, 875.0), (0.0273, 900.0), (0.0273, 1200.0)), 'dependencies': 0, 'type': ISOTROPIC}, 'name': 'Steel'}""")
                    mdb.models[mdbname].HomogeneousSolidSection(name='Steel', 
                        material='Steel', thickness=None)
                    mdb.models[mdbname].HomogeneousSolidSection(name='Fire resistance', 
                        material='Fire resistance material', thickness=None)
                    #Section assignment
                    p = mdb.models[mdbname].parts['steel']
                    f = p.faces
                    faces = f.getByBoundingBox()
                    region = p.Set(faces=faces, name='Face-1')
                    p.SectionAssignment(region=region, sectionName='Steel', offset=0.0, 
                        offsetType=MIDDLE_SURFACE, offsetField='', 
                        thicknessAssignment=FROM_SECTION)

                    p = mdb.models[mdbname].parts['Fire resistance']
                    f = p.faces
                    faces = f.getByBoundingBox()
                    region = p.Set(faces=faces, name='Face-2')
                    p.SectionAssignment(region=region, sectionName='Fire resistance', offset=0.0, 
                        offsetType=MIDDLE_SURFACE, offsetField='', 
                        thicknessAssignment=FROM_SECTION)



                    a = mdb.models[mdbname].rootAssembly
                    a.DatumCsysByDefault(CARTESIAN)
                    p = mdb.models[mdbname].parts['Fire resistance']
                    a.Instance(name='Fire resistance-1', part=p, dependent=OFF)
                    p = mdb.models[mdbname].parts['steel']
                    a.Instance(name='steel-1', part=p, dependent=OFF)
                    mdb.models[mdbname].HeatTransferStep(name='Step-1', previous='Initial', 
                        timePeriod=7200.0, maxNumInc=5000, initialInc=0.01, minInc=1e-16, 
                        maxInc=50.0, deltmx=10.0)
                    mdb.models[mdbname].FilmConditionProp(name='IntProp-1', 
                        temperatureDependency=OFF, dependencies=0, property=((2.5e-05, ), ))

                    mdb.models[mdbname].TabularAmplitude(name='Amp-1', timeSpan=STEP, 
                        smooth=SOLVER_DEFAULT, data=((0.0, 20.0), (60.0, 349.2136658), (120.0, 
                        444.5048779), (180.0, 502.289303), (240.0, 543.8873093), (300.0, 
                        576.4104306), (360.0, 603.1176476), (420.0, 625.7768252), (480.0, 
                        645.455108), (540.0, 662.8463867), (600.0, 678.4273315), (660.0, 
                        692.5395523), (720.0, 705.4362483), (780.0, 717.3103082), (840.0, 
                        728.312063), (900.0, 738.5609528), (960.0, 748.1534501), (1020.0, 
                        757.1685957), (1080.0, 765.6719608), (1140.0, 773.7185436), (1200.0, 
                        781.3549272), (1260.0, 788.6209131), (1320.0, 795.5507769), (1380.0, 
                        802.1742463), (1440.0, 808.5172716), (1500.0, 814.6026398), (1560.0, 
                        820.4504687), (1620.0, 826.0786082), (1680.0, 831.5029687), (1740.0, 
                        836.7377928), (1800.0, 841.7958797), (1860.0, 846.6887747), (1920.0, 
                        851.4269275), (1980.0, 856.0198265), (2040.0, 860.4761132), (2100.0, 
                        864.8036804), (2160.0, 869.0097558), (2220.0, 873.100975), (2280.0, 
                        877.0834446), (2340.0, 880.9627965), (2400.0, 884.7442362), (2460.0, 
                        888.4325848), (2520.0, 892.0323158), (2580.0, 895.5475878), (2640.0, 
                        898.9822734), (2700.0, 902.3399847), (2760.0, 905.6240963), (2820.0, 
                        908.8377658), (2880.0, 911.9839517), (2940.0, 915.0654299), (3000.0, 
                        918.0848086), (3060.0, 921.0445413), (3120.0, 923.946939), (3180.0, 
                        926.7941809), (3240.0, 929.5883242), (3300.0, 932.3313134), (3360.0, 
                        935.0249876), (3420.0, 937.671089), (3480.0, 940.2712687), (3540.0, 
                        942.8270936), (3600.0, 945.3400513), (3660.0, 947.8115564), (3720.0, 
                        950.2429541), (3780.0, 952.6355255), (3840.0, 954.990491), (3900.0, 
                        957.3090145), (3960.0, 959.5922069), (4020.0, 961.8411286), (4080.0, 
                        964.0567933), (4140.0, 966.2401703), (4200.0, 968.3921871), (4260.0, 
                        970.5137319), (4320.0, 972.6056555), (4380.0, 974.6687738), (4440.0, 
                        976.7038692), (4500.0, 978.7116928), (4560.0, 980.692966), (4620.0, 
                        982.6483816), (4680.0, 984.578606), (4740.0, 986.48428), (4800.0, 
                        988.3660202), (4860.0, 990.2244204), (4920.0, 992.0600525), (4980.0, 
                        993.8734676), (5040.0, 995.6651972), (5100.0, 997.4357536), (5160.0, 
                        999.1856316), (5220.0, 1000.915308), (5280.0, 1002.625245), (5340.0, 
                        1004.315888), (5400.0, 1005.987666), (5460.0, 1007.640997), (5520.0, 
                        1009.276283), (5580.0, 1010.893914), (5640.0, 1012.494267), (5700.0, 
                        1014.077707), (5760.0, 1015.644587), (5820.0, 1017.195251), (5880.0, 
                        1018.730032), (5940.0, 1020.24925), (6000.0, 1021.753218), (6060.0, 
                        1023.24224), (6120.0, 1024.71661), (6180.0, 1026.176612), (6240.0, 
                        1027.622525), (6300.0, 1029.054619), (6360.0, 1030.473153), (6420.0, 
                        1031.878384), (6480.0, 1033.270557), (6540.0, 1034.649914), (6600.0, 
                        1036.016688), (6660.0, 1037.371108), (6720.0, 1038.713393), (6780.0, 
                        1040.04376), (6840.0, 1041.362418), (6900.0, 1042.669572), (6960.0, 
                        1043.965421), (7020.0, 1045.250159), (7080.0, 1046.523974), (7140.0, 
                        1047.787051), (7200.0, 1049.039569)))



                    a = mdb.models[mdbname].rootAssembly
                    s1 = a.instances['Fire resistance-1'].edges

                    s2 = a.instances['steel-1'].edges
                    My_edges1=s1.getByBoundingBox(yMin=dep/2-1)
                    My_edges2=s1.getByBoundingBox(yMax=-(dep/2+insu-1))
                    My_edges3=s1.getByBoundingBox(xMax=-(wid/2+insu-1))
                    My_edges4=s1.getByBoundingBox(xMin=(wid/2+insu-1))
                    My_edges5=s1.getByBoundingBox(yMax=(dep/2-ft-insu),yMin=-(dep/2-insu-ft),xMax=-(wt/2+insu-1))
                    My_edges6=s1.getByBoundingBox(yMax=(dep/2-insu-ft),yMin=-(dep/2-insu-ft),xMin=(wt/2+insu-1))
                    side1Edges1 = s2.getByBoundingBox(yMin=dep/2-1)
                    side1Edges2 = My_edges1+My_edges2+My_edges3+My_edges4+My_edges5+My_edges6
                    a.Surface(side1Edges=side1Edges1+side1Edges2, name='Surf-3')
                    region=a.Surface(side1Edges=side1Edges1+side1Edges2, name='Surf-3')
                    mdb.models[mdbname].FilmCondition(name='Int-1', createStepName='Step-1', surface=region, definition=PROPERTY_REF, interactionProperty='IntProp-1', 
                        sinkTemperature=1.0, sinkAmplitude='Amp-1', sinkDistributionType=UNIFORM, sinkFieldName='')
                    mdb.models[mdbname].interactions['Int-1'].setValues(
                        definition=PROPERTY_REF, interactionProperty='IntProp-1', 
                        sinkTemperature=1.0, sinkAmplitude='Amp-1')
                    mdb.models[mdbname].RadiationToAmbient(name='Int-2', 
                        createStepName='Step-1', surface=region, radiationType=AMBIENT, 
                        distributionType=UNIFORM, field='', emissivity=0.3, ambientTemperature=1.0, 
                        ambientTemperatureAmp='Amp-1')
                    mdb.models[mdbname].fieldOutputRequests['F-Output-1'].setValues(
                        variables=('NT', 'TEMP', 'HFL', 'HFLA', 'RFL'))
                    mdb.models[mdbname].HistoryOutputRequest(name='H-Output-1', 
                        createStepName='Step-1', variables=('HFLA', 'HTL', 'HTLA', 'RADFL', 
                        'RADFLA', 'RADTL'))
                    mdb.Job(name='3sided'+mdbname, model=mdbname, description='', 
                        type=ANALYSIS, atTime=None, waitMinutes=0, waitHours=0, queue=None, 
                        memory=90, memoryUnits=PERCENTAGE, getMemoryFromAnalysis=True, 
                        explicitPrecision=SINGLE, nodalOutputPrecision=SINGLE, echoPrint=OFF, 
                        modelPrint=OFF, contactPrint=OFF, historyPrint=OFF, userSubroutine='', 
                        scratch='', resultsFormat=ODB, multiprocessingMode=DEFAULT, numCpus=1, 
                        numGPUs=1)

                    def Create_Partition_by_Plane(mdbname,id_plane):
                        a = mdb.models[mdbname].rootAssembly
                        f1=a.instances['steel-1'].faces[:]
                        f2=a.instances['Fire resistance-1'].faces[:]
                        f=f1+f2
                        d = a.datums
                        a.PartitionFaceByDatumPlane(faces=f,datumPlane=d[id_plane])

                    a = mdb.models[mdbname].rootAssembly
                    v1 = a.instances['Fire resistance-1'].vertices
                    dp1=a.DatumPlaneByTwoPoint(point1=v1[7], point2=v1[8])
                    id_plane=dp1.id

                    a = mdb.models[mdbname].rootAssembly
                    v11 = a.instances['Fire resistance-1'].vertices
                    dp2=a.DatumPlaneByTwoPoint(point1=v11[9], point2=v11[8])
                    id_plane=dp2.id



                    v1 = a.instances['steel-1'].vertices
                    dp3=a.DatumPlaneByTwoPoint(point1=v1[11], point2=v1[10])
                    id_plane=dp3.id


                    v11 = a.instances['Fire resistance-1'].vertices
                    dp4=a.DatumPlaneByTwoPoint(point1=v11[4], point2=v11[5])
                    id_plane=dp4.id

                    v1 = a.instances['steel-1'].vertices
                    dp5=a.DatumPlaneByTwoPoint(point1=v1[4], point2=v1[3])
                    id_plane=dp5.id

                    v11 = a.instances['steel-1'].vertices
                    dp6=a.DatumPlaneByTwoPoint(point1=v11[8], point2=v11[7])
                    id_plane=dp6.id

                    e1 = a.instances['Fire resistance-1'].edges
                    v1 = a.instances['Fire resistance-1'].vertices
                    dp7=a.DatumPlaneByTwoPoint(point2=v1[5], 
                        point1=a.instances['Fire resistance-1'].InterestingPoint(edge=e1[4], 
                        rule=MIDDLE))
                    id_plane=dp7.id

                    v11 = a.instances['Fire resistance-1'].vertices
                    e11 = a.instances['Fire resistance-1'].edges
                    dp8=a.DatumPlaneByTwoPoint(point1=v11[4], 
                        point2=a.instances['Fire resistance-1'].InterestingPoint(edge=e11[4], 
                        rule=MIDDLE))
                    id_plane=dp8.id

                    v1 = a.instances['Fire resistance-1'].vertices
                    e1 = a.instances['Fire resistance-1'].edges
                    dp9=a.DatumPlaneByPointNormal(point=v1[4], normal=e1[4])
                    id_plane=dp9.id

                    v11 = a.instances['Fire resistance-1'].vertices
                    e11 = a.instances['Fire resistance-1'].edges
                    dp10=a.DatumPlaneByPointNormal(point=v11[5], normal=e11[4])
                    id_plane=dp10.id

                    v1 = a.instances['steel-1'].vertices
                    e1 = a.instances['steel-1'].edges
                    dp11=a.DatumPlaneByPointNormal(point=v1[4], normal=e1[3])
                    id_plane=dp11.id

                    v11 = a.instances['steel-1'].vertices
                    e11 = a.instances['steel-1'].edges
                    dp12=a.DatumPlaneByPointNormal(point=v11[7], normal=e11[7])
                    id_plane=dp12.id



                    v1 = a.instances['steel-1'].vertices
                    e1 = a.instances['steel-1'].edges
                    dp13=a.DatumPlaneByPointNormal(point=v1[2], normal=e1[2])
                    id_plane=dp13.id


                    v11 = a.instances['steel-1'].vertices
                    e11 = a.instances['steel-1'].edges
                    dp14=a.DatumPlaneByPointNormal(point=v11[3], normal=e11[2])
                    id_plane=dp14.id

                    Create_Partition_by_Plane(mdbname,dp1.id)
                    Create_Partition_by_Plane(mdbname,dp2.id)
                    Create_Partition_by_Plane(mdbname,dp3.id)
                    Create_Partition_by_Plane(mdbname,dp4.id)
                    Create_Partition_by_Plane(mdbname,dp5.id)
                    Create_Partition_by_Plane(mdbname,dp6.id)
                    Create_Partition_by_Plane(mdbname,dp7.id)
                    Create_Partition_by_Plane(mdbname,dp8.id)
                    Create_Partition_by_Plane(mdbname,dp9.id)
                    Create_Partition_by_Plane(mdbname,dp10.id)
                    Create_Partition_by_Plane(mdbname,dp11.id)
                    Create_Partition_by_Plane(mdbname,dp12.id)
                    Create_Partition_by_Plane(mdbname,dp13.id)
                    Create_Partition_by_Plane(mdbname,dp14.id)




                    f1 = a.instances['Fire resistance-1'].faces
                    pickedface1=f1.getByBoundingBox()
                    f2 = a.instances['steel-1'].faces
                    pickedface2=f2.getByBoundingBox()
                    pickedFaces = pickedface1+pickedface2
                    f=pickedface1+pickedface2
                    session.viewports['Viewport: 1'].view.setValues(nearPlane=543.272, 
                        farPlane=1090.99, width=745.248, height=329.673, cameraPosition=(198.428, 
                        66.7408, 817.13), cameraTarget=(198.428, 66.7408, 0))
                    e=a.instances['steel-1'].edges
                    My_edges2=e.getByBoundingBox(yMax=-(dep/2-1))
                    My_edges3=e.getByBoundingBox(xMax=-(wid/2-1))
                    My_edges4=e.getByBoundingBox(xMin=(wid/2-1))
                    My_edges5=e.getByBoundingBox(yMax=(dep/2-ft),yMin=-(dep/2-ft),xMax=-(wt/2-1))
                    My_edges6=e.getByBoundingBox(yMax=(dep/2-ft),yMin=-(dep/2-ft),xMin=(wt/2-1))
                    My_edges=My_edges2+My_edges3+My_edges4+My_edges5+My_edges6
                    myRegion1=regionToolset.Region(edges=My_edges)
                    e=a.instances['Fire resistance-1'].edges
                    My_edges1=e.getByBoundingBox(yMin=-(dep/2-ft-1),yMax=(dep/2-ft-1),xMax=(wt/2+1),xMin=-(wt/2+1))
                    My_edges2=e.getByBoundingBox(yMin=(dep/2-ft-1),xMax=(wid/2+1),xMin=-(wid/2+1))
                    My_edges3=e.getByBoundingBox(yMax=-(dep/2-ft-1),yMin=-(dep/2+1),xMax=(wid/2+1),xMin=-(wid/2+1))
                    My_edges=My_edges1+My_edges2+My_edges3
                    myRegion2=regionToolset.Region(edges=My_edges)  
                    def Tie(mdbname,Tie_name,m_region,s_region):
                        mdb.models[mdbname].Tie(name=Tie_name,master=m_region,slave=s_region,positionTolerance=1e-1)
                    Tie(mdbname,'Tie-1',myRegion2,myRegion1)    
                    #datums=d1.keys()
                    #for i in datums:
                       # a.PartitionFaceByDatumPlane(faces=pickedface1,datumPlane=d1[9])





                    session.viewports['Viewport: 1'].view.setValues(nearPlane=502.108, 
                        farPlane=1132.15, width=1311.28, height=580.069, cameraPosition=(259.003, 
                        5.63809, 817.13), cameraTarget=(259.003, 5.63809, 0))

                    a = mdb.models[mdbname].rootAssembly
                    My_instance1=a.instances['Fire resistance-1']
                    My_instance2=a.instances['steel-1']
                    instances1=(My_instance1,My_instance2)
                    a.seedPartInstance(regions=instances1,size=1)
                    v1 = a.instances['steel-1'].vertices
                    v11 = a.instances['Fire resistance-1'].vertices
                    v=v1+v11
                    f1 = a.instances['Fire resistance-1'].faces
                    pickedface1=f1.getByBoundingBox()
                    f2 = a.instances['steel-1'].faces
                    pickedface2=f2.getByBoundingBox()
                    pickedFaces = pickedface1+pickedface2
                    f=pickedface1+pickedface2
                    My_region1=regionToolset.Region(faces=f)
                    v2=v1[:]+v11[:]
                    My_regions=(My_region1,)
                    def set_Mesh_Control(My_region):
                        a.setMeshControls(regions=My_region,elemShape=QUAD_DOMINATED,technique=STRUCTURED)
                    set_Mesh_Control(pickedFaces)
                    def generate_Mesh(mdbname):
                        a = mdb.models[mdbname].rootAssembly
                        My_instance1=a.instances['Fire resistance-1']
                        My_instance2=a.instances['steel-1']
                        My_instances=(My_instance1,My_instance2)
                        a.generateMesh(regions=My_instances)
                    generate_Mesh(mdbname)

                    My_instance1=a.instances['Fire resistance-1']
                    My_instance2=a.instances['steel-1']
                    e1=My_instance1.elements
                    e2=My_instance2.elements
                    My_meshes=(e1,e2)
                    Eset=a.Set(name='elset',elements=My_meshes)


                    elemType1 = mesh.ElemType(elemCode=DC2D4, elemLibrary=STANDARD)
                    elemType2 = mesh.ElemType(elemCode=DC2D3, elemLibrary=STANDARD)
                    f=(f1,f2)
                    def setElementType(mdbname,Eset,elemType1,elemType2):
                        a=mdb.models[mdbname].rootAssembly
                        a.setElementType(regions=Eset, elemTypes=(elemType1,elemType2))
                    setElementType(mdbname,f,elemType1,elemType2)

                    f1 = a.instances['Fire resistance-1'].faces
                    faces1 = f1.getByBoundingBox()
                    e1 = a.instances['Fire resistance-1'].edges
                    edges1 = e1.getByBoundingBox()
                    v1 = a.instances['Fire resistance-1'].vertices
                    verts1 = v1.getByBoundingBox()
                    f2 = a.instances['steel-1'].faces
                    faces2 = f2.getByBoundingBox()
                    e2 = a.instances['steel-1'].edges
                    edges2 = e2.getByBoundingBox()
                    v2 = a.instances['steel-1'].vertices
                    verts2 = v2.getByBoundingBox()
                    region = a.Set(vertices=verts1+verts2, edges=edges1+edges2, faces=faces1+faces2, name='Temp')
                    mdb.models[mdbname].Temperature(name='Predefined Field-1', 
                        createStepName='Initial', region=region, distributionType=UNIFORM, 
                        crossSectionDistribution=CONSTANT_THROUGH_THICKNESS, magnitudes=(20.0, ))



                    #My_region=regionToolset.Region(elements=My_meshes)


                    def Define_Upperflange_nodeSet (mdbname,set_name,y_Min):
                        n=mdb.models[mdbname].rootAssembly.instances['steel-1'].nodes
                        My_nodes1=n.getByBoundingBox(yMin=y_Min)
                        mdb.models[mdbname].rootAssembly.Set(name=set_name,nodes=My_nodes1)
                    def Define_Web_nodeSet (mdbname,set_name,y_Min,y_Max):
                        n=mdb.models[mdbname].rootAssembly.instances['steel-1'].nodes
                        My_nodes1=n.getByBoundingBox(yMin=y_Min,yMax=y_Max)
                        mdb.models[mdbname].rootAssembly.Set(name=set_name,nodes=My_nodes1)
                    def Define_Lowerflange_nodeSet (mdbname,set_name,y_Max):
                        n=mdb.models[mdbname].rootAssembly.instances['steel-1'].nodes
                        My_nodes1=n.getByBoundingBox(yMax=y_Max)
                        mdb.models[mdbname].rootAssembly.Set(name=set_name,nodes=My_nodes1)

                    #Define Nodes Set
                    y_Min=dep/2-ft-1
                    Define_Upperflange_nodeSet (mdbname,'Upperflange',y_Min)
                    y_Min=-(dep/2-ft-1)
                    y_Max=dep/2-ft-1
                    Define_Web_nodeSet (mdbname,'Web',y_Min,y_Max)
                    y_Max=-(dep/2-ft-1)
                    Define_Lowerflange_nodeSet (mdbname,'Lowerflange',y_Max)

                    y_Min1=dep/2-ft-0.5
                    y_Max1=dep/2-ft+0.5
                    y_Min2=-(dep/2-ft+0.5)
                    y_Max2=-(dep/2-ft-0.5)

                    #Get_Nodes_by_BoundingBox
                    def Get_Nodes_by_BoundingBox(My_mdb,My_instance,y_Min,y_Max): 
                        n=mdb.models[My_mdb].rootAssembly.instances[My_instance].nodes
                        My_nodes=n.getByBoundingBox(yMin=y_Min,yMax=y_Max)
                        return My_nodes
                    #Establish Node set
                    def establish_Set(mdbname,setname,My_nodes):
                        My_set=mdb.models[mdbname].rootAssembly.Set(name=setname,nodes=My_nodes)
                        return My_set

                    My_nodes=Get_Nodes_by_BoundingBox(mdbname,'steel-1',y_Min1,y_Max1)
                    establish_Set(mdbname,'topjunction',My_nodes)
                    My_nodes=Get_Nodes_by_BoundingBox(mdbname,'steel-1',y_Min2,y_Max2)
                    establish_Set(mdbname,'bottomjunction',My_nodes)








