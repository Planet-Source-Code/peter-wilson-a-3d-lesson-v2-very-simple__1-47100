MIDAR's Simple 3D
=================
This simple 3D application shows you just how simple 3D computer graphics can be to program. Pure VB code and the only mathematics is division. Simple hey?

Objects that are far away, appear smaller. Thus simply divide our 3D object's X and Y coordinates by it's Z coordinate.

  NewPixelX = X / Z
  NewPixelY = Y / Z

That's all you need for 3D computer graphics, and that's all this program does.


=========================
Keyboard & Mouse Controls
=========================
Move Camera Left/Right	=	Left and Right arrow keys.
Move Camera Up/Down	=	Up and Down arrow keys.
Move Camera In/Out	=	SHIFT-Up / SHIFT-Down

Zoom Camera In/Out	=	Page-Up / Page-Down
(Also known as 'Perspective Distortion' and/or 'Field Of View')

Mouse Move		=	Hi-lights dots that are close by	(optional routine - slows down program)

Reset Camera		=	Space Bar
Quit Application	=	Esc Key


=======
History
=======
Version 1.0 initial release to PSC on 22 July 03

Version 2.0 update to PSC on 24 July 03

	* In version 1.0, it wasn't immediately clear which way X, Y and Z pointed. In version 2.0 I have cleared this up considerably. Here are the updated coordinates:
	* Positive X points to the Right
	* Positive Z points *into* the monitor - away from You.
	* Positive Y goes Up

           +y
            |   +z (away from you - into the monitor)
            |  / 
            | /
	    |/
-x  --------+--------  +x
           /|
          / |
         /  |
       -z   |
	   -y

	Notice: The MS Operating Systems has Y=0 at the top of the monitor, with increasing values of Y going down towards the bottom. However, in this application I wanted the Origin (0,0,0) to be in the center of the screen, and for Y to go up so I flipped the sign from + to - (see subroutine: DoDisplay3DPoints)

	* Version two allows you to move the mouse, and have the dots hi-lighted that a close to the mouse. This is just for fun, and it also slows down the program... but I just wanted to show you how to do it.


Cheers,

Peter Wilson
http://dev.midar.com/
