Option Explicit

Public Const xDisplay As Integer = 26
Public Const yDisplay As Integer = 74

Public Const FirstHopperWall As Integer = 20
Public HopperWallHeight As Integer
Public Const HopperWallSep As Integer = 20

Public iMax As Integer
Public jMax As Integer
Public iPos As Integer
Public jPos As Integer

Public Hopletend1 As Integer            'location of end of hopperlet
Public Hopletend2 As Integer
Public Hopletend3 As Integer
Public Hopletend4 As Integer

Public h1Count As Integer               'hopperlet counter
Public h2Count As Integer
Public h3Count As Integer
Public h4Count As Integer

Public hb1Count As Integer              'hopper base counter
Public hb2Count As Integer
Public hb3Count As Integer
Public hb4Count As Integer

Public ht1Count As Integer              'hopper total counter
Public ht2Count As Integer
Public ht3Count As Integer
Public ht4Count As Integer

Public f1Count As Integer               'feeder counter
Public f2Count As Integer
Public f3Count As Integer
Public f4Count As Integer

Public height1 As Integer               'Height of product in hopper
Public height2 As Integer
Public height3 As Integer
Public height4 As Integer

Public Maxheight As Integer
Public Maxheight2 As Integer
Public Maxheight3 As Integer

Public f1 As Integer
Public f2 As Integer
Public f3 As Integer
Public f4 As Integer

Public Feedcount1 As Integer
Public Feedcount2 As Integer
Public Feedcount3 As Integer
Public Feedcount4 As Integer

Public Feedtotal As Integer

Public timeInc As Integer

Public Feedrate1 As Integer
Public Feedrate2 As Integer
Public Feedrate3 As Integer
Public Feedrate4 As Integer

Public startFeed1 As Integer
Public startFeed2 As Integer
Public startFeed3 As Integer
Public startFeed4 As Integer

Public iTrainPos As Integer

Public door1 As Single
Public door2 As Single
Public door3 As Single
Public door4 As Single

Public Offset As Integer
Public NewWagon As Integer

Public MaxFeedRate1 As Integer
Public MaxFeedRate2 As Integer
Public MaxFeedRate3 As Integer
Public MaxFeedRate4 As Integer


Public initialTime As Integer                  'used to calculate the time when each door opens
Public initialDistance As Integer               'used to door triggering positions and wagon location

Public TrigDist1 As Integer                    'location of triggers
Public TrigDist2 As Integer
Public TrigDist3 As Integer

Public Time As Integer
Public Blocks As Integer

Public dooropentime(4)

Public onewagononly As Integer
Public currentwagon As Integer

Public hcount(4)
Public massremoval(4)
Public massinfeeder(4)

Public roundingcounter1 As Single
Public roundingcounter2 As Single

Public blocksremoved1 As Single
Public blocksremoved2 As Single
Public blocksremoved3 As Single
Public blocksremoved4 As Single

Public bottomrow1 As Single
Public bottomrow2 As Single
Public bottomrow3 As Single
Public bottomrow4 As Single

Public HopperArray() As Boolean

