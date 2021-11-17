import femm
import math
from openpyxl import Workbook

# Magnetic couplings
#
#
#
# The objective of this program is to find the torque
# of the magnetic couplings. The diameter, steel width,
# magnet width and airgap and the number of poles can
# be changed. The error of it should be 0.0225% compared
# to the method of rotated. If I have some spare time,
# the rotated method should be completed soon. There is
# no limit of poles. The maximum number has been used is
# 100.
#
# Should be noticed:
# The number of poles should only be even numbers and
# more than 4 to avoid meaningless. The results are
# send into a excel which would named "result".
#
# Bugs:
# There are some bugs now for the 12 pole
# magnetic couplings. It cannot do the finite element
# analysis and cannot know the reason.
#
#


# Set the start and the end of number of poles wanted
num = 2
num_backup = num
end = 14

# Set the parameters for magnetic couplings
Defined_diameter = 150
Defined_diameter /= 2
air_gap = 2.0
smart_mesh = 1
poles = 22
degree = 360/poles

# Set the function of getting point
def Draw_Points(magnet_width):
	i = 0
	# Draw counterclockwise
	while i < poles:
		# Draw nodes
		femm.mi_addnode(diameter*math.cos(math.radians(degree*i)),diameter*math.sin(math.radians(degree*i)))
		femm.mi_addnode((diameter-magnet_width)*math.cos(math.radians(degree*i)),(diameter-magnet_width)*math.sin(math.radians(degree*i)))
		# Connect the nodes
		femm.mi_addsegment(diameter*math.cos(math.radians(degree*i)),diameter*math.sin(math.radians(degree*i)),(diameter-magnet_width)*math.cos(math.radians(degree*i)),(diameter-magnet_width)*math.sin(math.radians(degree*i)))
		if insidecircle == 1:
			femm.mi_selectsegment((diameter-magnet_width/2)*math.cos(math.radians(degree*i)),(diameter-magnet_width/2)*math.sin(math.radians(degree*i)))
			femm.mi_setgroup(1)
			femm.mi_clearselected()
		if i > 0:
			femm.mi_addarc(diameter*math.cos(math.radians(degree*(i-1))),diameter*math.sin(math.radians(degree*(i-1))),diameter*math.cos(math.radians(degree*i)),diameter*math.sin(math.radians(degree*i)),degree,1)
			femm.mi_addarc((diameter-magnet_width)*math.cos(math.radians(degree*(i-1))),(diameter-magnet_width)*math.sin(math.radians(degree*(i-1))),(diameter-magnet_width)*math.cos(math.radians(degree*i)),(diameter-magnet_width)*math.sin(math.radians(degree*i)),degree,1)
			if insidecircle == 1:
				femm.mi_selectarcsegment(diameter*math.cos(math.radians(degree*(i-1))),diameter*math.sin(math.radians(degree*(i-1))))
				femm.mi_selectarcsegment((diameter-magnet_width)*math.cos(math.radians(degree*(i-1))),(diameter-magnet_width)*math.sin(math.radians(degree*(i-1))))
				femm.mi_setgroup(1)
				femm.mi_clearselected()
		i+=1
	# Draw the last arc
	femm.mi_addarc(diameter*math.cos(math.radians(degree*(i-1))),diameter*math.sin(math.radians(degree*(i-1))),diameter*math.cos(math.radians(degree*i)),diameter*math.sin(math.radians(degree*i)),degree,1)
	femm.mi_addarc((diameter-magnet_width)*math.cos(math.radians(degree*(i-1))),(diameter-magnet_width)*math.sin(math.radians(degree*(i-1))),(diameter-magnet_width)*math.cos(math.radians(degree*i)),(diameter-magnet_width)*math.sin(math.radians(degree*i)),degree,1)
	if insidecircle == 1:
		femm.mi_selectarcsegment(diameter*math.cos(math.radians(degree*(i-1))),diameter*math.sin(math.radians(degree*(i-1))))
		femm.mi_selectarcsegment((diameter-magnet_width)*math.cos(math.radians(degree*(i-1))),(diameter-magnet_width)*math.sin(math.radians(degree*(i-1))))
		femm.mi_setgroup(1)
		femm.mi_clearselected()

# Set the block material
def Add_Material(magnet_width):
	t = 0.0
	while t < poles:
		if insidecircle == 0:
			femm.mi_addblocklabel((diameter-magnet_width/2)*math.cos(math.radians(degree*(0.5+t))),(diameter-magnet_width/2)*math.sin(math.radians(degree*(0.5+t))))
			femm.mi_selectlabel((diameter-magnet_width/2)*math.cos(math.radians(degree*(1/2+t))),(diameter-magnet_width/2)*math.sin(math.radians(degree*(1/2+t))))
		if insidecircle == 1:
			femm.mi_addblocklabel((diameter-magnet_width/2)*math.cos(math.radians(degree*t)),(diameter-magnet_width/2)*math.sin(math.radians(degree*t)))
			femm.mi_selectlabel((diameter-magnet_width/2)*math.cos(math.radians(degree*t)),(diameter-magnet_width/2)*math.sin(math.radians(degree*t)))
		if t%2 == 1:
			if insidecircle == 0:
				# femm.mi_setblockprop('NdFeB 32 MGOe',0,0.1,'<None>',(t+0.5)*degree-180,0,1)
				femm.mi_setblockprop('NdFeB 32 MGOe',1,0,'<None>',(t+0.5)*degree-180,0,1)
			if insidecircle == 1:
				# femm.mi_setblockprop('NdFeB 32 MGOe',0,0.1,'<None>',t*degree-180,0,1)
				femm.mi_setblockprop('NdFeB 32 MGOe',1,0,'<None>',t*degree-180,0,1)
		else:
			if insidecircle == 0:
				# femm.mi_setblockprop('NdFeB 32 MGOe',0,0.1,'<None>',((t+0.5)*degree)*pow(-1,t),0,1)
				femm.mi_setblockprop('NdFeB 32 MGOe',1,0,'<None>',((t+0.5)*degree)*pow(-1,t),0,1)
			if insidecircle == 1:
				# femm.mi_setblockprop('NdFeB 32 MGOe',0,0.1,'<None>',(t*degree)*pow(-1,t),0,1)
				femm.mi_setblockprop('NdFeB 32 MGOe',1,0,'<None>',(t*degree)*pow(-1,t),0,1)
		femm.mi_clearselected()
		t +=1.0


while num <= end:

	# Defined poles
	diameter = Defined_diameter
	magnet_width1 = num
	steel_width = 4

	# Start up and connect to FEMM
	femm.openfemm()

	# Create a new electrostatics problem
	femm.newdocument(0)


	# Flag to judge if its the inside circle
	insidecircle = 0

	#Flag to judge if the material is added for the outside circle
	materialadded = 0

	# Set a porblem
	femm.mi_probdef(0, 'millimeters','planar',1.e-8,10,30)

	# Open smart mesh
	if smart_mesh == 0:
		femm.mi_smartmesh(1)
		smart_mesh = 1

	if num == 6:
		femm.mi_smartmesh(0)
		smart_mesh = 0


	# Set the boundary conditions
	femm.mi_addboundprop('outside',0,0,0,0,0,0,0,0,0,0,0)

	# Set the materials
	femm.mi_getmaterial('Air')
	femm.mi_getmaterial('1006 Steel')
	femm.mi_getmaterial('NdFeB 32 MGOe')

	# Calculate area and make them equal
	magnet_area = (diameter - steel_width)*(diameter - steel_width)*math.pi-(diameter - steel_width-magnet_width1)*(diameter - steel_width-magnet_width1)*math.pi



	# Draw the greometry and add boundary condition
	femm.mi_drawarc(diameter,0,-1*diameter,0,180,1)
	femm.mi_addarc(-1*diameter,0,diameter,0,180,1)
	femm.mi_selectarcsegment(0,diameter)
	femm.mi_selectarcsegment(0,-1*diameter)
	femm.mi_setarcsegmentprop(1,'outside',0,0)
	femm.mi_clearselected()

	# Add material property for the air outside
	femm.mi_addblocklabel(diameter,diameter)
	femm.mi_selectlabel(diameter,diameter)
	# femm.mi_setblockprop('Air',0,0.1,'<None>',0,0,1)
	femm.mi_setblockprop('Air',1,0,'<None>',0,0,1)
	femm.mi_clearselected()

	# Add material property for the metal of the large circle
	femm.mi_addblocklabel(diameter-steel_width/2,0)
	femm.mi_selectlabel(diameter-steel_width/2,0)
	# femm.mi_setblockprop('1006 Steel',0,0.1,'<None>',0,0,1)
	femm.mi_setblockprop('1006 Steel',1,0,'<None>',0,0,1)
	femm.mi_clearselected()

	# Draw the magnets and add material properties
	diameter -= steel_width
	Draw_Points(magnet_width1)
	Add_Material(magnet_width1)

	# Add material property for the metal and of the air gap
	femm.mi_addblocklabel(diameter-(air_gap/2+magnet_width1),0)
	femm.mi_selectlabel(diameter-(air_gap/2+magnet_width1),0)
	# femm.mi_setblockprop('Air',0,0.1,'<None>',0,0,1)
	femm.mi_setblockprop('Air',1,0,'<None>',0,0,1)
	femm.mi_clearselected()

	magnet_width2 = ((diameter-magnet_width1-air_gap)*(diameter-magnet_width1-air_gap) - magnet_area/math.pi)**0.5

	# For inside circle
	diameter -= (air_gap + magnet_width1)
	insidecircle = 1
	magnet_width2 = diameter - magnet_width2
	Draw_Points(magnet_width2)
	femm.mi_selectgroup(1)
	femm.mi_moverotate(0,0,degree/2)
	femm.mi_clearselected()
	Add_Material(magnet_width2)

	# Add material property for the metal of the small circle
	femm.mi_addblocklabel(diameter/2,0)
	femm.mi_selectlabel(diameter/2,0)
	# femm.mi_setblockprop('1006 Steel',0,0.1,'<None>',0,0,1)
	femm.mi_setblockprop('1006 Steel',1,0,'<None>',0,0,1)
	femm.mi_clearselected()

	# # Save the geometry to disk so we can analyze it
	femm.mi_saveas('strips.FEM')

	# Now,analyze the problem and load the solution when the analysis is finished
	femm.mi_analyze()
	femm.mi_loadsolution()

	# Get the result
	result = 0.0
	m = 0
	while m < poles:
		femm.mo_selectblock((diameter-magnet_width2/2)*math.cos(math.radians(degree*m)),(diameter-magnet_width2/2)*math.sin(math.radians(degree*m)))
		m += 1
	femm.mo_selectblock(diameter/2,0)
	result = femm.mo_blockintegral(22)
	print('result is %g\n' % result)

	femm.closefemm()

	# Write it into workbook
	if num == num_backup:
		mybook = Workbook()
	wa = mybook.active
	middle = num - 1
	String = 'A'+ str(middle)
	wa[String] = result
	mybook.save('result.xlsx')

	# Making it a loop
	num += 1






