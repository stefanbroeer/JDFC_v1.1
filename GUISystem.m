function varargout = GUISystem(varargin)
% GUISYSTEM MATLAB code for GUISystem.fig
%      GUISYSTEM, by itself, creates a new GUISYSTEM or raises the existing
%      singleton*.
%
%      H = GUISYSTEM returns the handle to a new GUISYSTEM or the handle to
%      the existing singleton*.
%
%      GUISYSTEM('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in GUISYSTEM.M with the given input arguments.
%
%      GUISYSTEM('Property','Value',...) creates a new GUISYSTEM or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before GUISystem_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to GUISystem_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help GUISystem

% Last Modified by GUIDE v2.5 18-Nov-2019 10:00:58

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @GUISystem_OpeningFcn, ...
                   'gui_OutputFcn',  @GUISystem_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT
clear
clc


% --- Executes just before GUISystem is made visible.
function GUISystem_OpeningFcn(hObject, ~, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to GUISystem (see VARARGIN)

% Choose default command line output for GUISystem
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes GUISystem wait for user response (see UIRESUME)
% uiwait(handles.figure1);
% digits(10)
%Create tab group

handles.tgroup = uitabgroup('Parent', handles.figure1,'TabLocation', 'left');
handles.tab0 = uitab('Parent', handles.tgroup, 'Title', 'PreSets');
handles.tab1 = uitab('Parent', handles.tgroup, 'Title', 'Input');
handles.tab2 = uitab('Parent', handles.tgroup, 'Title', 'Ala/Cys');
handles.tab3 = uitab('Parent', handles.tgroup, 'Title', 'Asp/Glu');
handles.tab4 = uitab('Parent', handles.tgroup, 'Title', 'Phe/Gly');
handles.tab5 = uitab('Parent', handles.tgroup, 'Title', 'His/IsoL');
handles.tab6 = uitab('Parent', handles.tgroup, 'Title', 'Lys/Leu');
handles.tab7 = uitab('Parent', handles.tgroup, 'Title', 'Met/Asn');
handles.tab8 = uitab('Parent', handles.tgroup, 'Title', 'Pro/Gln');
handles.tab9 = uitab('Parent', handles.tgroup, 'Title', 'Arg/Ser');
handles.tab10 = uitab('Parent', handles.tgroup, 'Title', 'Thr/Val');
handles.tab11 = uitab('Parent', handles.tgroup, 'Title', 'Trp/Tyr');
handles.tab12 = uitab('Parent', handles.tgroup, 'Title', 'AAT Flux');
%Place panels into each tab
set(handles.p0,'Parent',handles.tab0)
set(handles.p1,'Parent',handles.tab1)
set(handles.p2,'Parent',handles.tab2)
set(handles.p3,'Parent',handles.tab3)
set(handles.p4,'Parent',handles.tab4)
set(handles.p5,'Parent',handles.tab5)
set(handles.p6,'Parent',handles.tab6)
set(handles.p7,'Parent',handles.tab7)
set(handles.p8,'Parent',handles.tab8)
set(handles.p9,'Parent',handles.tab9)
set(handles.p10,'Parent',handles.tab10)
set(handles.p11,'Parent',handles.tab11)
set(handles.p20,'Parent',handles.tab12)

%Reposition each panel to same location as panel 1  
set(handles.p0,'position',get(handles.p1,'position'));
set(handles.p2,'position',get(handles.p1,'position'));
set(handles.p3,'position',get(handles.p1,'position'));
set(handles.p4,'position',get(handles.p1,'position'));
set(handles.p5,'position',get(handles.p1,'position'));
set(handles.p6,'position',get(handles.p1,'position'));
set(handles.p7,'position',get(handles.p1,'position'));
set(handles.p8,'position',get(handles.p1,'position'));
set(handles.p9,'position',get(handles.p1,'position'));
set(handles.p10,'position',get(handles.p1,'position'));
set(handles.p11,'position',get(handles.p1,'position'));
set(handles.p20,'position',get(handles.p1,'position'));


% --- Outputs from this function are returned to the command line.
function varargout = GUISystem_OutputFcn(~, ~, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in Transport.
function Transport_Callback(~, ~, handles) %#ok<*DEFNU>
% hObject    handle to Transport (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%--------------------------------------------------------------------------%
% Call data

cla reset
tic
set(handles.Indicator,'BackgroundColor',[0.6 0.04 0.15]);
drawnow

iterations=str2double(get(handles.iterations,'string'));
AAfile=strcat(matlab.lang.makeValidName(get(handles.AA,'string')),'.xlsx');
Kminfile=strcat(matlab.lang.makeValidName(get(handles.Km,'string')),'.xlsx');
Kmexfile=strcat(matlab.lang.makeValidName(get(handles.Km,'string')),'.xlsx');
Vmaxfile=strcat(matlab.lang.makeValidName(get(handles.Vmax,'string')),'.xlsx');
rVmaxfile=strcat(matlab.lang.makeValidName(get(handles.rVmax,'string')),'.xlsx');
AAdepfile=strcat(matlab.lang.makeValidName(get(handles.AAdep,'string')),'.xlsx');
AAconfile=strcat(matlab.lang.makeValidName(get(handles.AAcon,'string')),'.xlsx');

Kmin=table2array(readtable(Kminfile,'range','B55:W102','ReadVariableNames',false));
Kmex=table2array(readtable(Kmexfile,'range','B3:W50','ReadVariableNames',false));
Vmax=table2array(readtable(Vmaxfile,'range','B3:B50','ReadVariableNames',false));
rVmax=table2array(readtable(rVmaxfile,'range','B3:W50','ReadVariableNames',false));
AAcon=table2array(readtable(AAconfile,'range','D4:Y25','ReadVariableNames',false));
AAdep=table2array(readtable(AAdepfile,'range','B2:B23','ReadVariableNames',false));

%------------------------------------------------------------------------%
%Define Variables

%SET INTRA AND EXTRA AA CONCENTRATION AS MATRIX
AA=readtable(AAfile);
AAex=table2array(AA(1,2:23));
AA0in=table2array(AA(2,2:23));
AAin=zeros(iterations,22);
AAin(1,:)=AA0in;

%DETERMINE ION CONCENTRATION AND KM VALUES
Nain=table2array(AA(2,24));
Kpin=table2array(AA(2,25));
Hyin=table2array(AA(2,26));
Clin=table2array(AA(2,27));
Naex=table2array(AA(1,24));
Kpex=table2array(AA(1,25));
Hyex=table2array(AA(1,26));
Clex=table2array(AA(1,27));

NaKmin=table2array(readtable(Kminfile,'range','X55:X102','ReadVariableNames',false));
NaKmex=table2array(readtable(Kmexfile,'range','X3:X50','ReadVariableNames',false));
HyKmin=table2array(readtable(Kminfile,'range','Y55:Y102','ReadVariableNames',false));
HyKmex=table2array(readtable(Kmexfile,'range','Y3:Y50','ReadVariableNames',false));
ClKmin=table2array(readtable(Kminfile,'range','AA55:AA102','ReadVariableNames',false));
ClKmex=table2array(readtable(Kmexfile,'range','AA3:AA50','ReadVariableNames',false));
KpKmin=table2array(readtable(Kminfile,'range','Z55:Z102','ReadVariableNames',false));
KpKmex=table2array(readtable(Kmexfile,'range','Z3:Z50','ReadVariableNames',false));

%DEFINE TRANSPORTER VMAX
EAAT3Vmax=Vmax(1);
EAAT1Vmax=Vmax(2);
EAAT2Vmax=Vmax(3);
ASCT2Vmax=Vmax(5);
ASCT1Vmax=Vmax(4);
EAAT4Vmax=Vmax(6);
EAAT5Vmax=Vmax(7);

GlyT2Vmax=Vmax(9);
GlyT1Vmax=Vmax(10);
PROTVmax=Vmax(11);
ATB0Vmax=Vmax(12);
B0AT2Vmax=Vmax(13);
NTT4Vmax=Vmax(14);
XT2Vmax=Vmax(15);
B0AT1Vmax=Vmax(16);
SIT1Vmax=Vmax(17);

CAT1Vmax=Vmax(19);
CAT2Vmax=Vmax(20);
CAT3Vmax=Vmax(21);
CAT4Vmax=Vmax(22);
LAT1Vmax=Vmax(23);
yLAT2Vmax=Vmax(24);
yLAT1Vmax=Vmax(25);
LAT2Vmax=Vmax(26);
b0ATVmax=Vmax(27);
Asc1Vmax=Vmax(28);
xCTVmax=Vmax(29);
Asc2Vmax=Vmax(30);
AGT1Vmax=Vmax(31);

TAT1Vmax=Vmax(33);

PAT1Vmax=Vmax(35);
PAT2Vmax=Vmax(36);
PAT4Vmax=Vmax(37);

SNAT1Vmax=Vmax(39);
SNAT2Vmax=Vmax(40);
SNAT3Vmax=Vmax(41);
SNAT4Vmax=Vmax(42);
SNAT5Vmax=Vmax(43);
SNAT7Vmax=Vmax(44);
SNAT8Vmax=Vmax(45);

LAT3Vmax=Vmax(47);
LAT4Vmax=Vmax(48);


%DEFINE CONSTANTS
Vol=str2double(get(handles.Vol,'String'));
dt=str2double(get(handles.dt,'String'));
time=1:iterations;
psi=-0.06;
z1=1;
z2=2;
F=96500;
R=8.314;
T=310.15;
B1=exp((-z1*F*0.5*psi)/(R*T));
B2=exp((-z2*F*0.5*psi)/(R*T));

%ENABLE ACTIVATION OF EACH CLICKBUTTON
EAAT3=get(handles.EAAT3,'Value');
EAAT1=get(handles.EAAT1,'Value');
EAAT2=get(handles.EAAT2,'Value');
ASCT2=get(handles.ASCT2,'Value');
ASCT1=get(handles.ASCT1,'Value');
EAAT4=get(handles.EAAT4,'Value');
EAAT5=get(handles.EAAT5,'Value');
GlyT2=get(handles.GLyT2,'Value');
GlyT1=get(handles.GlyT1,'Value');
PROT=get(handles.PROT,'Value');
ATB0=get(handles.ATB0,'Value');
B0AT2=get(handles.B0AT2,'Value');
NTT4=get(handles.NTT4,'Value');
XT2=get(handles.XT2,'Value');
B0AT1=get(handles.B0AT1,'Value');
SIT1=get(handles.SIT1,'Value');
LAT1=get(handles.LAT1,'Value');
LAT2=get(handles.LAT2,'Value');
yLAT1=get(handles.yLAT1,'Value');
yLAT2=get(handles.yLAT2,'Value');
xCT=get(handles.xCT,'Value');
CAT1=get(handles.CAT1,'Value');
CAT2=get(handles.CAT2,'Value');
CAT3=get(handles.CAT3,'Value');
CAT4=get(handles.CAT4,'Value');
b0AT=get(handles.b0AT,'Value');
Asc1=get(handles.Asc1,'Value');
Asc2=get(handles.Asc2,'Value');
AGT1=get(handles.AGT1,'Value');
PAT1=get(handles.PAT1,'Value');
PAT2=get(handles.PAT2,'Value');
PAT4=get(handles.PAT4,'Value');
TAT1=get(handles.TAT1,'Value');
SNAT1=get(handles.SNAT1,'Value');
SNAT2=get(handles.SNAT2,'Value');
SNAT3=get(handles.SNAT3,'Value');
SNAT4=get(handles.SNAT4,'Value');
SNAT5=get(handles.SNAT5,'Value');
SNAT7=get(handles.SNAT7,'Value');
SNAT8=get(handles.SNAT8,'Value');
LAT3=get(handles.LAT3,'Value');
LAT4=get(handles.LAT4,'Value');

% CALL RELATIVE VMAX FOR EACH TRANSPORTER
%DEFINE TRANSPORTER VMAX
EAAT3rVmax=rVmax(1,:);
EAAT1rVmax=rVmax(2,:);
EAAT2rVmax=rVmax(3,:);
ASCT2rVmax=rVmax(5,:);
ASCT1rVmax=rVmax(4,:);
EAAT4rVmax=rVmax(6,:);
EAAT5rVmax=rVmax(7,:);

GlyT2rVmax=rVmax(9,:);
GlyT1rVmax=rVmax(10,:);
PROTrVmax=rVmax(11,:);
ATB0rVmax=rVmax(12,:);
B0AT2rVmax=rVmax(13,:);
NTT4rVmax=rVmax(14,:);
XT2rVmax=rVmax(15,:);
B0AT1rVmax=rVmax(16,:);
SIT1rVmax=rVmax(17,:);

CAT1rVmax=rVmax(19,:);
CAT2rVmax=rVmax(20,:);
CAT3rVmax=rVmax(21,:);
CAT4rVmax=rVmax(22,:);
LAT1rVmax=rVmax(23,:);
yLAT2rVmax=rVmax(24,:);
yLAT1rVmax=rVmax(25,:);
LAT2rVmax=rVmax(26,:);
b0ATrVmax=rVmax(27,:);
Asc1rVmax=rVmax(28,:);
xCTrVmax=rVmax(29,:);
Asc2rVmax=rVmax(30,:);
AGT1rVmax=rVmax(31,:);

TAT1rVmax=rVmax(33,:);

PAT1rVmax=rVmax(35,:);
PAT2rVmax=rVmax(36,:);
PAT4rVmax=rVmax(37,:);

SNAT1rVmax=rVmax(39,:);
SNAT2rVmax=rVmax(40,:);
SNAT3rVmax=rVmax(41,:);
SNAT4rVmax=rVmax(42,:);
SNAT5rVmax=rVmax(43,:);
SNAT7rVmax=rVmax(44,:);
SNAT8rVmax=rVmax(45,:);

LAT3rVmax=rVmax(47,:);
LAT4rVmax=rVmax(48,:);


%EXTRACELLULAR KM VALUES
EAAT3Kmex=Kmex(1,:);
EAAT1Kmex=Kmex(2,:);
EAAT2Kmex=Kmex(3,:);
ASCT2Kmex=Kmex(5,:);
ASCT1Kmex=Kmex(4,:);
EAAT4Kmex=Kmex(6,:);
EAAT5Kmex=Kmex(7,:);

GlyT2Kmex=Kmex(9,:);
GlyT1Kmex=Kmex(10,:);
PROTKmex=Kmex(11,:);
ATB0Kmex=Kmex(12,:);
B0AT2Kmex=Kmex(13,:);
NTT4Kmex=Kmex(14,:);
XT2Kmex=Kmex(15,:);
B0AT1Kmex=Kmex(16,:);
SIT1Kmex=Kmex(17,:);

CAT1Kmex=Kmex(19,:);
CAT2Kmex=Kmex(20,:);
CAT3Kmex=Kmex(21,:);
CAT4Kmex=Kmex(22,:);
LAT1Kmex=Kmex(23,:);
yLAT2Kmex=Kmex(24,:);
yLAT1Kmex=Kmex(25,:);
LAT2Kmex=Kmex(26,:);
b0ATKmex=Kmex(27,:);
Asc1Kmex=Kmex(28,:);
xCTKmex=Kmex(29,:);
Asc2Kmex=Kmex(30,:);
AGT1Kmex=Kmex(31,:);

TAT1Kmex=Kmex(33,:);

PAT1Kmex=Kmex(35,:);
PAT2Kmex=Kmex(36,:);
PAT4Kmex=Kmex(37,:);

SNAT1Kmex=Kmex(39,:);
SNAT2Kmex=Kmex(40,:);
SNAT3Kmex=Kmex(41,:);
SNAT4Kmex=Kmex(42,:);
SNAT5Kmex=Kmex(43,:);
SNAT7Kmex=Kmex(44,:);
SNAT8Kmex=Kmex(45,:);

LAT3Kmex=Kmex(47,:);
LAT4Kmex=Kmex(48,:);

%INTRACELLULAR KM VALUES
EAAT3Kmin=Kmin(1,:);
EAAT1Kmin=Kmin(2,:);
EAAT2Kmin=Kmin(3,:);
ASCT2Kmin=Kmin(5,:);
ASCT1Kmin=Kmin(4,:);
EAAT4Kmin=Kmin(6,:);
EAAT5Kmin=Kmin(7,:);

GlyT2Kmin=Kmin(9,:);
GlyT1Kmin=Kmin(10,:);
PROTKmin=Kmin(11,:);
ATB0Kmin=Kmin(12,:);
B0AT2Kmin=Kmin(13,:);
NTT4Kmin=Kmin(14,:);
XT2Kmin=Kmin(15,:);
B0AT1Kmin=Kmin(16,:);
SIT1Kmin=Kmin(17,:);

CAT1Kmin=Kmin(19,:);
CAT2Kmin=Kmin(20,:);
CAT3Kmin=Kmin(21,:);
CAT4Kmin=Kmin(22,:);
LAT1Kmin=Kmin(23,:);
yLAT2Kmin=Kmin(24,:);
yLAT1Kmin=Kmin(25,:);
LAT2Kmin=Kmin(26,:);
b0ATKmin=Kmin(27,:);
Asc1Kmin=Kmin(28,:);
xCTKmin=Kmin(29,:);
Asc2Kmin=Kmin(30,:);
AGT1Kmin=Kmin(31,:);

TAT1Kmin=Kmin(33,:);

PAT1Kmin=Kmin(35,:);
PAT2Kmin=Kmin(36,:);
PAT4Kmin=Kmin(37,:);

SNAT1Kmin=Kmin(39,:);
SNAT2Kmin=Kmin(40,:);
SNAT3Kmin=Kmin(41,:);
SNAT4Kmin=Kmin(42,:);
SNAT5Kmin=Kmin(43,:);
SNAT7Kmin=Kmin(44,:);
SNAT8Kmin=Kmin(45,:);

LAT3Kmin=Kmin(47,:);
LAT4Kmin=Kmin(48,:);


% Predefine Matrices

dEAAT3indt=zeros(iterations,22);
dEAAT1indt=zeros(iterations,22);
dEAAT2indt=zeros(iterations,22);
dASCT2indt=zeros(iterations,22);
dASCT1indt=zeros(iterations,22);
dEAAT4indt=zeros(iterations,22);
dEAAT5indt=zeros(iterations,22);

dGlyT2indt=zeros(iterations,22);
dGlyT1indt=zeros(iterations,22);
dPROTindt=zeros(iterations,22);
dATB0indt=zeros(iterations,22);
dB0AT2indt=zeros(iterations,22);
dNTT4indt=zeros(iterations,22);
dXT2indt=zeros(iterations,22);
dB0AT1indt=zeros(iterations,22);
dSIT1indt=zeros(iterations,22);

dCAT1indt=zeros(iterations,22);
dCAT2indt=zeros(iterations,22);
dCAT3indt=zeros(iterations,22);
dCAT4indt=zeros(iterations,22);
dLAT1indt=zeros(iterations,22);
dyLAT2indt=zeros(iterations,22);
dyLAT1indt=zeros(iterations,22);
dLAT2indt=zeros(iterations,22);
db0ATindt=zeros(iterations,22);
dAsc1indt=zeros(iterations,22);
dxCTindt=zeros(iterations,22);
dAsc2indt=zeros(iterations,22);
dAGT1indt=zeros(iterations,22);

dTAT1indt=zeros(iterations,22);

dPAT1indt=zeros(iterations,22);
dPAT2indt=zeros(iterations,22);
dPAT4indt=zeros(iterations,22);

dSNAT1indt=zeros(iterations,22);
dSNAT2indt=zeros(iterations,22);
dSNAT3indt=zeros(iterations,22);
dSNAT4indt=zeros(iterations,22);
dSNAT5indt=zeros(iterations,22);
dSNAT7indt=zeros(iterations,22);
dSNAT8indt=zeros(iterations,22);
dLAT3indt=zeros(iterations,22);
dLAT4indt=zeros(iterations,22);

%------------------------------------------------------------------------%
% RUN THROUGH EACH FUNCTION FOR AMINO ACID CONCENTRATION DETERMINATION


for i=2:iterations
    
    if EAAT3
        [dEAAT3indt(i,:)]=ComplexSym1_function(AAex,AAin(i-1,:),EAAT3Kmex,EAAT3Kmin,EAAT3rVmax,EAAT3Vmax,NaKmin(1),NaKmex(1),KpKmin(1),KpKmex(1),HyKmin(1),HyKmex(1),Naex,Nain,Kpin,Kpex,Hyin,Hyex,dt,Vol,B2);
    else
        [dEAAT3indt(i,22)]=deal(0);
    end
    if EAAT1
        [dEAAT1indt(i,:)]=ComplexSym1_function(AAex,AAin(i-1,:),EAAT1Kmex,EAAT1Kmin,EAAT1rVmax,EAAT1Vmax,NaKmin(2),NaKmex(2),KpKmin(2),KpKmex(2),HyKmin(2),HyKmex(2),Naex,Nain,Kpin,Kpex,Hyin,Hyex,dt,Vol,B2);
    else
        [dEAAT1indt(i,22)]=deal(0);
    end
    if EAAT2
        [dEAAT2indt(i,:)]=ComplexSym1_function(AAex,AAin(i-1,:),EAAT2Kmex,EAAT2Kmin,EAAT2rVmax,EAAT2Vmax,NaKmin(3),NaKmex(3),KpKmin(3),KpKmex(3),HyKmin(3),HyKmex(3),Naex,Nain,Kpin,Kpex,Hyin,Hyex,dt,Vol,B2);
    else
        [dEAAT2indt(i,22)]=deal(0);
    end
    if ASCT2
        [dASCT2indt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),ASCT2Kmex,ASCT2Kmin,ASCT2rVmax,ASCT2Vmax,dt,Vol);
    else
        [dASCT2indt(i,22)]=deal(0);
    end
    if ASCT1
        [dASCT1indt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),ASCT1Kmex,ASCT1Kmin,ASCT1rVmax,ASCT1Vmax,dt,Vol);
    else
        [dASCT1indt(i,22)]=deal(0);
    end
    if EAAT4
        [dEAAT4indt(i,:)]=ComplexSym1_function(AAex,AAin(i-1,:),EAAT4Kmex,EAAT4Kmin,EAAT4rVmax,EAAT4Vmax,NaKmin(6),NaKmex(6),KpKmin(6),KpKmex(6),HyKmin(6),HyKmex(6),Naex,Nain,Kpin,Kpex,Hyin,Hyex,dt,Vol,B2);
    else
        [dEAAT4indt(i,22)]=deal(0);
    end
    if EAAT5
        [dEAAT5indt(i,:)]=ComplexSym1_function(AAex,AAin(i-1,:),EAAT5Kmex,EAAT5Kmin,EAAT5rVmax,EAAT5Vmax,NaKmin(7),NaKmex(7),KpKmin(7),KpKmex(7),HyKmin(7),HyKmex(7),Naex,Nain,Kpin,Kpex,Hyin,Hyex,dt,Vol,B2);
    else
        [dEAAT5indt(i,22)]=deal(0);
    end
    
    if GlyT2
        [dGlyT2indt(i,:)]=ComplexSym3_function(AAex,AAin(i-1,:),GlyT2Kmex,GlyT2Kmin,GlyT2rVmax,GlyT2Vmax,NaKmin(9),NaKmex(9),ClKmin(9),ClKmex(9),Naex,Nain,Clin,Clex,dt,Vol,B2);
    else
        [dGlyT2indt(i,22)]=deal(0);
    end
    if GlyT1
        [dGlyT1indt(i,:)]=ComplexSym2_function(AAex,AAin(i-1,:),GlyT1Kmex,GlyT1Kmin,GlyT1rVmax,GlyT1Vmax,NaKmin(10),NaKmex(10),ClKmin(10),ClKmex(10),Naex,Nain,Clin,Clex,dt,Vol,B1);
    else
        [dGlyT1indt(i,22)]=deal(0);
    end
    if PROT
        [dPROTindt(i,:)]=NaSym_function(AAex,AAin(i-1,:),PROTKmex,PROTKmin,PROTrVmax,PROTVmax,NaKmex(11),NaKmin(11),Naex,Nain,dt,Vol,B1);
    else
        [dPROTindt(i,22)]=deal(0);
    end
    if ATB0
        [dATB0indt(i,:)]=ComplexSym2_function(AAex,AAin(i-1,:),ATB0Kmex,ATB0Kmin,ATB0rVmax,ATB0Vmax,NaKmin(12),NaKmex(12),ClKmin(12),ClKmex(12),Naex,Nain,Clin,Clex,dt,Vol,B1);
    else
        [dATB0indt(i,22)]=deal(0);
    end
    if B0AT2
        [dB0AT2indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),B0AT2Kmex,B0AT2Kmin,B0AT2rVmax,B0AT2Vmax,NaKmex(13),NaKmin(13),Naex,Nain,dt,Vol,B1);
    else
        [dB0AT2indt(i,22)]=deal(0);
    end
    if NTT4
        [dNTT4indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),NTT4Kmex,NTT4Kmin,NTT4rVmax,NTT4Vmax,NaKmex(14),NaKmin(14),Naex,Nain,dt,Vol,B1);
    else
        [dNTT4indt(i,22)]=deal(0);
    end
    if XT2
        [dXT2indt(i,:)]=ComplexSym2_function(AAex,AAin(i-1,:),XT2Kmex,XT2Kmin,XT2rVmax,XT2Vmax,NaKmin(15),NaKmex(15),ClKmin(15),ClKmex(15),Naex,Nain,Clin,Clex,dt,Vol,B1);
    else
        [dXT2indt(i,22)]=deal(0);
    end
    if B0AT1
        [dB0AT1indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),B0AT1Kmex,B0AT1Kmin,B0AT1rVmax,B0AT1Vmax,NaKmex(16),NaKmin(16),Naex,Nain,dt,Vol,B1);
    else
        [dB0AT1indt(i,22)]=deal(0);
    end
    if SIT1
        [dSIT1indt(i,:)]=ComplexSym2_function(AAex,AAin(i-1,:),SIT1Kmex,SIT1Kmin,SIT1rVmax,SIT1Vmax,NaKmin(17),NaKmex(17),ClKmin(17),ClKmex(17),Naex,Nain,Clin,Clex,dt,Vol,B1);
    else
        [dSIT1indt(i,22)]=deal(0);
    end
    
    if CAT1
        [dCAT1indt(i,:)]=CatUniporter_function(AAex,AAin(i-1,:),CAT1Kmex,CAT1Kmin,CAT1rVmax,CAT1Vmax,dt,Vol,B1);
    else
        [dCAT1indt(i,22)]=deal(0);
    end
    if CAT2
        [dCAT2indt(i,:)]=CatUniporter_function(AAex,AAin(i-1,:),CAT2Kmex,CAT2Kmin,CAT2rVmax,CAT2Vmax,dt,Vol,B1);
    else
        [dCAT2indt(i,22)]=deal(0);
    end
    if CAT3
        [dCAT3indt(i,:)]=CatUniporter_function(AAex,AAin(i-1,:),CAT3Kmex,CAT3Kmin,CAT3rVmax,CAT3Vmax,dt,Vol,B1);
    else
        [dCAT3indt(i,22)]=deal(0);
    end
    if CAT4
        [dCAT4indt(i,:)]=CatUniporter_function(AAex,AAin(i-1,:),CAT4Kmex,CAT4Kmin,CAT4rVmax,CAT4Vmax,dt,Vol,B1);
    else
        [dCAT4indt(i,22)]=deal(0);
    end
    if LAT1
        [dLAT1indt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),LAT1Kmex,LAT1Kmin,LAT1rVmax,LAT1Vmax,dt,Vol);
    else
        [dLAT1indt(i,22)]=deal(0);
    end
    if yLAT2
        [dyLAT2indt(i,:)]=ComplexAntiporter_function(AAex,AAin(i-1,:),yLAT2Kmex,yLAT2Kmin,yLAT2rVmax,yLAT2Vmax,NaKmex(24),NaKmin(24),Naex,Nain,dt,Vol);
    else
        [dyLAT2indt(i,22)]=deal(0);
    end
    if yLAT1
        [dyLAT1indt(i,:)]=ComplexAntiporter_function(AAex,AAin(i-1,:),yLAT1Kmex,yLAT1Kmin,yLAT1rVmax,yLAT1Vmax,NaKmex(25),NaKmin(25),Naex,Nain,dt,Vol);
    else
        [dyLAT1indt(i,22)]=deal(0);
    end
    if LAT2
        [dLAT2indt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),LAT2Kmex,LAT2Kmin,LAT2rVmax,LAT2Vmax,dt,Vol);
    else
        [dLAT2indt(i,22)]=deal(0);
    end
    if b0AT
        [db0ATindt(i,:)]=ElectrogenicAntiporter(AAex,AAin(i-1,:),b0ATKmex,b0ATKmin,b0ATrVmax,b0ATVmax,dt,Vol,B1);
    else
        [db0ATindt(i,22)]=deal(0);
    end
    if Asc1
        [dAsc1indt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),Asc1Kmex,Asc1Kmin,Asc1rVmax,Asc1Vmax,dt,Vol);
    else
        [dAsc1indt(i,22)]=deal(0);
    end
    if xCT
        [dxCTindt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),xCTKmex,xCTKmin,xCTrVmax,xCTVmax,dt,Vol);
    else
        [dxCTindt(i,22)]=deal(0);
    end
    if Asc2
        [dAsc2indt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),Asc2Kmex,Asc2Kmin,Asc2rVmax,Asc2Vmax,dt,Vol);
    else
        [dAsc2indt(i,22)]=deal(0);
    end
    if AGT1
        [dAGT1indt(i,:)]=Antiporter_function(AAex,AAin(i-1,:),AGT1Kmex,AGT1Kmin,AGT1rVmax,AGT1Vmax,dt,Vol);
    else
        [dAGT1indt(i,22)]=deal(0);
    end
    
    if SNAT1
        [dSNAT1indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),SNAT1Kmex,SNAT1Kmin,SNAT1rVmax,SNAT1Vmax,NaKmex(39),NaKmin(39),Naex,Nain,dt,Vol,B1);
    else
        [dSNAT1indt(i,22)]=deal(0);
    end
    if SNAT2
        [dSNAT2indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),SNAT2Kmex,SNAT2Kmin,SNAT2rVmax,SNAT2Vmax,NaKmex(40),NaKmin(40),Naex,Nain,dt,Vol,B1);
    else
        [dSNAT2indt(i,22)]=deal(0);
    end
    if SNAT3
        [dSNAT3indt(i,:)]=ElectroSym_function(AAex,AAin(i-1,:),SNAT3Kmex,SNAT3Kmin,SNAT3rVmax,SNAT3Vmax,NaKmex(41),NaKmin(41),HyKmin(41),HyKmex(41),Naex,Nain,Hyex,Hyin,dt,Vol);
    else
        [dSNAT3indt(i,22)]=deal(0);
    end
    if SNAT4
        [dSNAT4indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),SNAT4Kmex,SNAT4Kmin,SNAT4rVmax,SNAT4Vmax,NaKmex(42),NaKmin(42),Naex,Nain,dt,Vol,B1);
    else
        [dSNAT4indt(i,22)]=deal(0);
    end
    if SNAT5
        [dSNAT5indt(i,:)]=ElectroSym_function(AAex,AAin(i-1,:),SNAT5Kmex,SNAT5Kmin,SNAT5rVmax,SNAT5Vmax,NaKmex(43),NaKmin(43),HyKmin(43),HyKmex(43),Naex,Nain,Hyex,Hyin,dt,Vol);
    else
        [dSNAT5indt(i,22)]=deal(0);
    end
    if SNAT7
        [dSNAT7indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),SNAT7Kmex,SNAT7Kmin,SNAT7rVmax,SNAT7Vmax,NaKmex(44),NaKmin(44),Naex,Nain,dt,Vol,B1);
    else
        [dSNAT7indt(i,22)]=deal(0);
    end
    if SNAT8
        [dSNAT8indt(i,:)]=NaSym_function(AAex,AAin(i-1,:),SNAT8Kmex,SNAT8Kmin,SNAT8rVmax,SNAT8Vmax,NaKmex(45),NaKmin(45),Naex,Nain,dt,Vol,B1);
    else
        [dSNAT8indt(i,22)]=deal(0);
    end
    
    if TAT1
        [dTAT1indt(i,:)]=Uniporter_function(AAex,AAin(i-1,:),TAT1Kmex,TAT1Kmin,TAT1rVmax,TAT1Vmax,dt,Vol);
    else
        [dTAT1indt(i,22)]=deal(0);
    end
    
    if PAT1
        [dPAT1indt(i,:)]=ProtonSym_function(AAex,AAin(i-1,:),PAT1Kmex,PAT1Kmin,PAT1rVmax,PAT1Vmax,HyKmin(35),HyKmex(35),Hyex,Hyin,dt,Vol,B1);
    else
        [dPAT1indt(i,22)]=deal(0);
    end
    if PAT2
        [dPAT2indt(i,:)]=ProtonSym_function(AAex,AAin(i-1,:),PAT2Kmex,PAT2Kmin,PAT2rVmax,PAT2Vmax,HyKmin(36),HyKmex(36),Hyex,Hyin,dt,Vol,B1);
    else
        [dPAT2indt(i,22)]=deal(0);
    end
    if PAT4
        [dPAT4indt(i,:)]=ProtonSym_function(AAex,AAin(i-1,:),PAT4Kmex,PAT4Kmin,PAT4rVmax,PAT4Vmax,HyKmin(37),HyKmex(37),Hyex,Hyin,dt,Vol,B1);
    else
        [dPAT4indt(i,22)]=deal(0);
    end
    
    if LAT3
        [dLAT3indt(i,:)]=Uniporter_function(AAex,AAin(i-1,:),LAT3Kmex,LAT3Kmin,LAT3rVmax,LAT3Vmax,dt,Vol);
    else
        [dLAT3indt(i,22)]=deal(0);
    end
    if LAT4
        [dLAT4indt(i,:)]=Uniporter_function(AAex,AAin(i-1,:),LAT4Kmex,LAT4Kmin,LAT4rVmax,LAT4Vmax,dt,Vol);
    else
        [dLAT4indt(i,22)]=deal(0);
    end
    
    
    %------------------------------------------------------------------------%
    %DETERMINE INTRACELLULAR CONCENTRATION GIVEN THE CHANGE IN CONCENTRATION THROUGH EACH TRANSPORTER
    
    
    AAin(i,:)=dEAAT3indt(i,:)+dEAAT1indt(i,:)+dEAAT2indt(i,:)+dASCT2indt(i,:)+dASCT1indt(i,:)+dEAAT4indt(i,:)+dEAAT5indt(i,:)+dGlyT2indt(i,:)+dGlyT1indt(i,:)+dPROTindt(i,:)+dATB0indt(i,:)+dB0AT2indt(i,:)+dNTT4indt(i,:)+dXT2indt(i,:)+dB0AT1indt(i,:)+dSIT1indt(i,:)+dCAT1indt(i,:)+dCAT2indt(i,:)+dCAT3indt(i,:)+dCAT4indt(i,:)+dLAT1indt(i,:)+dyLAT2indt(i,:)+dyLAT1indt(i,:)+dLAT2indt(i,:)+db0ATindt(i,:)+dAsc1indt(i,:)+dxCTindt(i,:)+dAsc2indt(i,:)+dAGT1indt(i,:)+dTAT1indt(i,:)+dPAT1indt(i,:)+dPAT2indt(i,:)+dPAT4indt(i,:)+dSNAT1indt(i,:)+dSNAT2indt(i,:)+dSNAT3indt(i,:)+dSNAT4indt(i,:)+dSNAT5indt(i,:)+dSNAT7indt(i,:)+dSNAT8indt(i,:)+dLAT3indt(i,:)+dLAT4indt(i,:)+AAin(i-1,:);
    
    % AMINO ACID CONVERSION AND DEPLETION
    % where n=AA being synthesised, m=AA being depleted
    
    conversion=zeros(22,22);
    
    for n=1:22
        for m=1:22
        conversion(n,m)=AAcon(n,m)*AAin(i,m);
        end
    end
    
    for n=1:22
    usage=sum(conversion);
    synthesis=sum(conversion,2);
    AAin(i,n)=(AAin(i,n)*(1-AAdep(n)))-usage(n)+synthesis(n);
    end
    
end

%------------------------------------------------------------------------%
%PLOT AA CONCENTRATION

cla(handles.axes1)
axes(handles.axes1)
hold on;

dAAindt=zeros(1,22);

for i=1:22
    dAAindt(i)=dEAAT3indt(2,i)+dEAAT1indt(2,i)+dEAAT2indt(2,i)+dASCT2indt(2,i)+dASCT1indt(2,i)+dEAAT4indt(2,i)+dEAAT5indt(2,i)+dGlyT2indt(2,i)+dGlyT1indt(2,i)+dPROTindt(2,i)+dATB0indt(2,i)+dB0AT2indt(2,i)+dNTT4indt(2,i)+dXT2indt(2,i)+dB0AT1indt(2,i)+dSIT1indt(2,i)+dCAT1indt(2,i)+dCAT2indt(2,i)+dCAT3indt(2,i)+dCAT4indt(2,i)+dLAT1indt(2,i)+dyLAT2indt(2,i)+dyLAT1indt(2,i)+dLAT2indt(2,i)+db0ATindt(2,i)+dAsc1indt(2,i)+dxCTindt(2,i)+dAsc2indt(2,i)+dTAT1indt(2,i)+dPAT1indt(2,i)+dPAT2indt(2,i)+dPAT4indt(2,i)+dSNAT1indt(2,i)+dSNAT2indt(2,i)+dSNAT3indt(2,i)+dSNAT4indt(2,i)+dSNAT5indt(2,i)+dSNAT7indt(2,i)+dSNAT8indt(2,i)+dLAT3indt(2,i)+dLAT4indt(2,i);
end

h=zeros(1,22);
x=gobjects(1,22);
AACategory={{AAin(:,1),'#B03A2E','Ala'},{AAin(:,2),'#AF7AC5','Cys'},{AAin(:,3),'#5499C7','Asp'},{AAin(:,4),'#48C9B0','Glu'},{AAin(:,5),'#F1C40F','Phe'},{AAin(:,6),'#EB984E','Gly'},{AAin(:,7),'#99A3A4','His'},{AAin(:,8),'#2E4053','Iso'},{AAin(:,9),'#52BE80','Lys'},{AAin(:,10),'#00FFFF','Leu'},{AAin(:,11),'#922B21','Met'},{AAin(:,12),'#ABEBC6','Asn'},{AAin(:,13),'#CCFF66','Orn'},{AAin(:,14),'#FF6600','Pro'},{AAin(:,15),'#0000FF','Gln'},{AAin(:,16),'#CC0033','Arg'},{AAin(:,17),'#CACFD2','Ser'},{AAin(:,18),'#66CC00','Thr'},{AAin(:,19),'#9933CC','Val'},{AAin(:,20),'#FF6699','Trp'},{AAin(:,21),'#CC99CC','Tyr'},{AAin(:,22),'000000','Csn'}};

for i=1:22
    vector=AACategory{i}{1};
    colour=AACategory{i}{2};
    if dAAindt(i)~=0
        x(i)=plot(time,vector,'color',hex2rgb(colour));
        h(i)=1;
    else
        x(i)=0;
        h(i)=0;
    end
end

hold off;

title('Amino Acid Flux','fontsize',12)
xlabel('Iterations','fontsize',8)
ylabel('Intracellular Concentration (然)','fontsize',8)

%-------------------------------------------------------------------------%
%DISPLAY ALANINE FLUX
cla(handles.axes2)
axes(handles.axes2);
hold on;

dAATindt=cat(3,dEAAT3indt,dEAAT1indt,dEAAT2indt,dASCT2indt,dASCT1indt,dEAAT4indt,dEAAT5indt,dGlyT2indt,dGlyT1indt,dPROTindt,dATB0indt,dB0AT2indt,dNTT4indt,dXT2indt,dB0AT1indt,dSIT1indt,dCAT1indt,dCAT2indt,dCAT3indt,dCAT4indt,dLAT1indt,dyLAT2indt,dyLAT1indt,dLAT2indt,db0ATindt,dAsc1indt,dxCTindt,dAsc2indt,dAGT1indt,dTAT1indt,dPAT1indt,dPAT2indt,dPAT4indt,dSNAT1indt,dSNAT2indt,dSNAT3indt,dSNAT4indt,dSNAT5indt,dSNAT7indt,dSNAT8indt,dLAT3indt,dLAT4indt);
AATCategory={{'EAAT3','#008000'},{'EAAT1','#FA8072'},{'EAAT2','#00FFFF'},{'ASCT2','#0000FF'},{'ASCT1','#008080'},{'EAAT4','#B0E0E6'},{'EAAT5','#EE82EE'},{'GlyT2','95A5A6'},{'GlyT1','#FFE4E1'},{'PROT','#FF5733'},{'ATB0','#FFDEAD'},{'B0AT2','#A52A2A'},{'NTT4','#BC8F8F'},{'XT2','#D2B48C'},{'B0AT1','#7FFFD4'},{'SIT1','#000000'},{'CAT1','#808080'},{'CAT2','#FF00FF'},{'CAT3','#F0E68C'},{'CAT4','#FFD700'},{'LAT1','#808000'},{'yLAT2','#ABEBC6'},{'yLAT1','00FF00'},{'LAT2','#0099CC'},{'b0AT','#B3FFCC'},{'Asc1','#F08080'},{'xCT','#FF0000'},{'Asc2','#D8BFD8'},{'AGT1','#FFEBCD'},{'TAT1','#F5E766'},{'PAT1','#9A7D0A'},{'PAT2','#800000'},{'PAT4','000000'},{'SNAT1','#800080'},{'SNAT2','#1F618D'},{'SNAT3','#B3B8DB'},{'SNAT4','#E74C3C'},{'SNAT5','#A9CCE3'},{'SNAT7','#8FFFB8'},{'SNAT8','#FFF566'},{'LAT3','#229954'},{'LAT4','#D35400'}};
totaltransporters=size(dAATindt,3);
alastring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,1,i)~=0
        plot(time,dAATindt(:,1,i),'color',hex2rgb(AATcolour));
        alastring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        alastring(i)={[]};
    end
end
set(handles.alalist,'string',alastring)
hold off;
title('Change in Alanine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)


%------------------------------------------------------------------------%
%DISPLAY CYSTEINE FLUX
cla(handles.axes3)
axes(handles.axes3);
hold on;
cysstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,2,i)~=0
        plot(time,dAATindt(:,2,i),'color',hex2rgb(AATcolour));
        cysstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        cysstring(i)={[]};
    end
end
hold off;
title('Change in Cysteine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.cyslist,'string',cysstring)

%--------------------------------------------------------------------------%
%DISPLAY ASPARTATE FLUX

cla(handles.aspaxes)
axes(handles.aspaxes);
hold on;
aspstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,3,i)~=0
        plot(time,dAATindt(:,3,i),'color',hex2rgb(AATcolour));
        aspstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        aspstring(i)={[]};
    end
end

hold off;

title('Change in Asparagine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.asplist,'string',aspstring)

%------------------------------------------------------------------------%
%DISPLAY GLUTAMATE FLUX

cla(handles.gluaxes)
axes(handles.gluaxes);
hold on;
glustring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,4,i)~=0
        plot(time,dAATindt(:,4,i),'color',hex2rgb(AATcolour));
        glustring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        glustring(i)={[]};
    end
end

hold off;
title('Change in Glutamate flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.glulist,'string',glustring)

%------------------------------------------------------------------------%
%DISPLAY PHENYLALANINE FLUX

cla(handles.pheaxes)
axes(handles.pheaxes);
hold on;
phestring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,5,i)~=0
        plot(time,dAATindt(:,5,i),'color',hex2rgb(AATcolour));
        phestring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        phestring(i)={[]};
    end
end

hold off;
title('Change in Phenylalanine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.phelist,'string',phestring)

%------------------------------------------------------------------------%
%DISPLAY GLYCINE FLUX

cla(handles.glyaxes)
axes(handles.glyaxes);
hold on;
glystring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,6,i)~=0
        plot(time,dAATindt(:,6,i),'color',hex2rgb(AATcolour));
        glystring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        glystring(i)={[]};
    end
end

hold off;
title('Change in Glycine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.glylist,'string',glystring)

%------------------------------------------------------------------------%
%DISPLAY HISTIDINE FLUX

cla(handles.hisaxes)
axes(handles.hisaxes);
hold on;
hisstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,7,i)~=0
        plot(time,dAATindt(:,7,i),'color',hex2rgb(AATcolour));
        hisstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        hisstring(i)={[]};
    end
end

hold off;
title('Change in Histidine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.hislist,'string',hisstring)

%------------------------------------------------------------------------%
%DISPLAY ISOLEUCINE FLUX

cla(handles.isoaxes)
axes(handles.isoaxes);
hold on;
isostring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,8,i)~=0
        plot(time,dAATindt(:,8,i),'color',hex2rgb(AATcolour));
        isostring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        isostring(i)={[]};
    end
end

hold off;
title('Change in Isoleucine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.isolist,'string',isostring)

%------------------------------------------------------------------------%
%DISPLAY LYSINE FLUX

cla(handles.lysaxes)
axes(handles.lysaxes);
hold on;
lysstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,9,i)~=0
        plot(time,dAATindt(:,9,i),'color',hex2rgb(AATcolour));
        lysstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        lysstring(i)={[]};
    end
end

hold off;
title('Change in Lysine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.lyslist,'string',lysstring)

%------------------------------------------------------------------------%
%DISPLAY LEUCINE FLUX

cla(handles.leuaxes)
axes(handles.leuaxes);
hold on;
leustring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,10,i)~=0
        plot(time,dAATindt(:,10,i),'color',hex2rgb(AATcolour));
        leustring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        leustring(i)={[]};
    end
end

hold off;
title('Change in Leucine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.leulist,'string',leustring)

%------------------------------------------------------------------------%
%DISPLAY METHIONINE FLUX

cla(handles.metaxes)
axes(handles.metaxes);
hold on;
metstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,11,i)~=0
        plot(time,dAATindt(:,11,i),'color',hex2rgb(AATcolour));
        metstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        metstring(i)={[]};
    end
end

hold off;
title('Change in Methionine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.metlist,'string',metstring)

%------------------------------------------------------------------------%
%DISPLAY ASPARAGINE FLUX

cla(handles.asnaxes)
axes(handles.asnaxes);
hold on;
asnstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,12,i)~=0
        plot(time,dAATindt(:,12,i),'color',hex2rgb(AATcolour));
        asnstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        asnstring(i)={[]};
    end
end

hold off;
title('Change in Asparagine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.asnlist,'string',asnstring)

%------------------------------------------------------------------------%
%DISPLAY PROLINE FLUX

cla(handles.proaxes)
axes(handles.proaxes);
hold on;
prostring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,14,i)~=0
        plot(time,dAATindt(:,14,i),'color',hex2rgb(AATcolour));
        prostring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        prostring(i)={[]};
    end
end

hold off;
title('Change in Proline flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.prolist,'string',prostring)

%------------------------------------------------------------------------%
%DISPLAY GLUTAMINE FLUX

cla(handles.glnaxes)
axes(handles.glnaxes);
hold on;
glnstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,15,i)~=0
        plot(time,dAATindt(:,15,i),'color',hex2rgb(AATcolour));
        glnstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        glnstring(i)={[]};
    end
end

hold off;
title('Change in Glutamine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.glnlist,'string',glnstring)

%------------------------------------------------------------------------%
%DISPLAY ARGININE FLUX

cla(handles.argaxes)
axes(handles.argaxes);
hold on;
argstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,16,i)~=0
        plot(time,dAATindt(:,16,i),'color',hex2rgb(AATcolour));
        argstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        argstring(i)={[]};
    end
end

hold off;
title('Change in Arginine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.arglist,'string',argstring)

%------------------------------------------------------------------------%
%DISPLAY SERINE FLUX

cla(handles.seraxes)
axes(handles.seraxes);
hold on;
serstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,17,i)~=0
        plot(time,dAATindt(:,17,i),'color',hex2rgb(AATcolour));
        serstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        serstring(i)={[]};
    end
end

hold off;
title('Change in Serine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.serlist,'string',serstring)

%------------------------------------------------------------------------%
%DISPLAY THREONINE FLUX

cla(handles.thraxes)
axes(handles.thraxes);
hold on;
thrstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,18,i)~=0
        plot(time,dAATindt(:,18,i),'color',hex2rgb(AATcolour));
        thrstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        thrstring(i)={[]};
    end
end

hold off;
title('Change in Threonine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.thrlist,'string',thrstring)

%------------------------------------------------------------------------%
%DISPLAY VALINE FLUX

cla(handles.valaxes)
axes(handles.valaxes);
hold on;
valstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,19,i)~=0
        plot(time,dAATindt(:,19,i),'color',hex2rgb(AATcolour));
        valstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        valstring(i)={[]};
    end
end
hold off;
title('Change in Valine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.vallist,'string',valstring)

%------------------------------------------------------------------------%
%DISPLAY TRYPTOPHAN FLUX

cla(handles.trpaxes)
axes(handles.trpaxes);
hold on;
trpstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,20,i)~=0
        plot(time,dAATindt(:,20,i),'color',hex2rgb(AATcolour));
        trpstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        trpstring(i)={[]};
    end
end

hold off;
title('Change in Tryptophan flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.trplist,'string',trpstring)

%------------------------------------------------------------------------%
%DISPLAY TYROSINE FLUX

cla(handles.tyraxes)
axes(handles.tyraxes);
hold on;
tyrstring=cell(1,totaltransporters);
for i=1:41
    AATname=AATCategory{i}{1};
    AATcolour=AATCategory{i}{2};
    if dAATindt(iterations,21,i)~=0
        plot(time,dAATindt(:,21,i),'color',hex2rgb(AATcolour));
        tyrstring(i)={['<HTML><font color=',AATcolour,'>',AATname,'</font>']};
    else
        tyrstring(i)={[]};
    end
end

hold off;
title('Change in Tyrosine flux','fontsize',12)
xlabel('Iterations(i)','fontsize',8)
ylabel('Change in Cytosolic Conc(然)','fontsize',8)
set(handles.tyrlist,'string',tyrstring)

%----------------------------------------------------------------------%
% AA FLUX THROUGH SINGLE TRANSPORTER DSIPLAY

cla(handles.AATaxes)
axes(handles.AATaxes);
hold on;
p=zeros(totaltransporters,22);
m=gobjects(totaltransporters,22);
for n=1:40
    for i=1:22
        colour=AACategory{i}{2};
        if dAATindt(iterations,i,n)~=0
            m(n,i)=plot(time,dAATindt(:,i,n),'color',hex2rgb(colour));
            p(n,i)=1;
        else
            m(n,i)=0;
            p(n,i)=0;
        end
    end
end
hold off;
title('AA Flux Through Single Transporter','fontsize',12)
ylabel('Change in Cytosolic AA Conc (\Delta然/dt)','fontsize',10)
xlabel('Iteration','fontsize',10)

e=toc;
set(handles.CPU,'string',e)
set(handles.Indicator,'BackgroundColor',[0.47 0.67 0.19]);


% Make intracellular amino acid concetrations global

setappdata(handles.Transport,'x',x);
setappdata(handles.Transport,'AATCategory',AATCategory);
setappdata(handles.Transport,'AACategory',AACategory);
setappdata(handles.Transport,'h',h);
setappdata(handles.Transport,'m',m);
setappdata(handles.Transport,'p',p);
setappdata(handles.Transport,'AAin',AAin);
setappdata(handles.Transport,'dAATindt',dAATindt);
setappdata(handles.Transport,'totaltransporters',totaltransporters);

% --- Executes on selection change in listbox.
function listbox_Callback(hObject, eventdata, handles) %#ok<INUSL>
% hObject    handle to listbox (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns listbox contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox


contents = cellstr(get(handles.listbox,'string'));

h=getappdata(handles.Transport,'h');
x=getappdata(handles.Transport,'x');
AAin=getappdata(handles.Transport,'AAin');
AACategory=getappdata(handles.Transport,'AACategory');
iterations=str2double(get(handles.iterations,'string'));

for i=1:22
    if h(i)==1
        set(x(i),'visible','off')
    end
end

for i=1:22
    name=AACategory{i}{3};
    colour=AACategory{i}{2};
    if contents{get(handles.listbox,'Value')}== name
        fAAin=num2str(AAin(iterations,i));
        set(handles.conc,'string',fAAin,'ForegroundColor',hex2rgb(colour));
        if h(i)==1
            set(x(i),'visible','on')
        end
    end
end

 % --- Executes on button press in resetAAT.
function resetAAT_Callback(~, ~, handles)
% hObject    handle to resetAAT (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
m=getappdata(handles.Transport,'m');
p=getappdata(handles.Transport,'p');

for i=1:22
    for n=1:41
        if p(n,i)==1
            set(m(n,i),'visible','on')
        end
    end
end

% --- Executes on button press in export.
function export_Callback(~, ~, handles)
% hObject    handle to export (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

exportfolder=matlab.lang.makeValidName(get(handles.exportfolder,'string'));
iterations=str2double(get(handles.iterations,'string'));
step=str2double(get(handles.step,'string'));
range=str2double(get(handles.range,'string'));
header={'Ala','Cys','Asp','Glu','Phe','Gly','His','IsoL','Lys','Leu','Met','Asn','Orn','Pro','Gln','Arg','Ser','Thr','Val','Trp','Tyr','Csn'};
c_header=transpose(header);

AAin=num2cell(getappdata(handles.Transport,'AAin'));
AA=num2cell(cell2mat(transpose(AAin(iterations,:))));

if isnan(step)||isnan(range)
    AAvector=AAin;
else
    AAvector=AAin(1:step:range,1:end);
end

X=horzcat(c_header,AA);
Y=vertcat(header,AAvector);
warning( 'off', 'MATLAB:xlswrite:AddSheet' ) 
xlswrite(strcat('C:/',exportfolder,'/AminoAcidConcentration.xls'),Y,2)
xlswrite(strcat('C:/',exportfolder,'/AminoAcidConcentration.xls'),X,1)


% --- Executes on button press in reset.
function reset_Callback(~, ~, handles)
% hObject    handle to reset (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

x=getappdata(handles.Transport,'x');
h=getappdata(handles.Transport,'h');

for i=1:22
    if h(i)==1
        set(x(i),'visible','on')
    end
end

% --- Executes on selection change in transporters.
function transporters_Callback(~, ~, handles)
% hObject    handle to transporters (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns transporters contents as cell array
%        contents{get(hObject,'Value')} returns selected item from transporters

contents = cellstr(get(handles.transporters,'string'));
m=getappdata(handles.Transport,'m');
p=getappdata(handles.Transport,'p');
AATCategory=getappdata(handles.Transport,'AATCategory');
totaltransporters=getappdata(handles.Transport,'totaltransporters');

for n=1:totaltransporters
    for i=1:22
        if p(n,i)==1
            set(m(n,i),'visible','off')
        end
    end
end

for n=1:totaltransporters
    for i=1:22
        name=AATCategory{n}{1};
        if isequal(contents{get(handles.transporters,'Value')},name)
            if p(n,i)==1
                set(m(n,i),'visible','on')
            end
        end
    end
end

% --- Executes on selection change in AATAAlist.
function AATAAlist_Callback(~, ~, handles)
% hObject    handle to AATAAlist (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns AATAAlist contents as cell array
%        contents{get(hObject,'Value')} returns selected item from AATAAlist

AAcontents = cellstr(get(handles.AATAAlist,'string'));
AATcontents = cellstr(get(handles.transporters,'string'));
p=getappdata(handles.Transport,'p');
m=getappdata(handles.Transport,'m');
AACategory=getappdata(handles.Transport,'AACategory');
AATCategory=getappdata(handles.Transport,'AATCategory');
totaltransporters=getappdata(handles.Transport,'totaltransporters');

for n=1:totaltransporters
    for i=1:22
        if p(n,i)==1
            set(m(n,i),'visible','off')
        end
    end
end

for i=1:22
    for n=1:totaltransporters
        AAname=AACategory{i}{3};
        AATname=AATCategory{n}{1};
        if isequal(AATcontents{get(handles.transporters,'Value')},AATname)
            if isequal(AAcontents{get(handles.AATAAlist,'Value')},AAname)
                if p(n,i)==1
                set(m(n,i),'visible','on')
                end
            end
        end
    end
end

function iterations_Callback(~, ~, ~)
function iterations_CreateFcn(hObject, ~, ~) 
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
%function LAT3_Callback(~, ~, ~)
function listbox_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function SNAT1_Callback(~, ~, ~)
function CAT1_Callback(~, ~, ~)
function LAT1_Callback(~, ~, ~)
function ASCT2_Callback(~, ~, ~)
function ASCT1_Callback(~, ~, ~)
function CAT2_Callback(~, ~, ~)
function EAAT1_Callback(~, ~, ~)
function EAAT2_Callback(~, ~, ~)
function EAAT3_Callback(~, ~, ~)
function GLyT2_Callback(~, ~, ~)
function LAT2_Callback(~, ~, ~)
function LAT4_Callback(~, ~, ~)
function LAT3_Callback(~, ~, ~)
function PAT1_Callback(~, ~, ~)
function PAT2_Callback(~, ~, ~)
function PAT4_Callback(~, ~, ~)
function SNAT2_Callback(~, ~, ~)
function SNAT4_Callback(~, ~, ~)
function SNAT3_Callback(~, ~, ~)
function xCT_Callback(~, ~, ~)
function yLAT1_Callback(~, ~, ~)
function yLAT2_Callback(~, ~, ~)
function alalist_CellEditCallback(~, ~, ~)
function alalist_Callback(~, ~, ~)
function alalist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function prolist_Callback(~, ~, ~)
function prolist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function glnlist_Callback(~, ~, ~)
function glnlist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function cyslist_Callback(~, ~, ~)
function cyslist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function glylist_Callback(~, ~, ~)
function glylist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function phelist_Callback(~, ~, ~)
function phelist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function hislist_Callback(~, ~, ~)
function hislist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function leulist_Callback(~, ~, ~)
function leulist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function lyslist_Callback(~, ~, ~)
function lyslist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function asnlist_Callback(~, ~, ~)
function asnlist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function metlist_Callback(~, ~, ~)
function metlist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function serlist_Callback(~, ~, ~)
function serlist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function arglist_Callback(~, ~, ~)
function arglist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function vallist_Callback(~, ~, ~)
function vallist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function thrlist_Callback(~, ~, ~)
function thrlist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function tyrlist_Callback(~, ~, ~)
function tyrlist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function trplist_Callback(~, ~, ~)
function trplist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function glulist_Callback(~, ~, ~)
function glulist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function asplist_Callback(~, ~, ~)
function asplist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function isolist_Callback(~, ~, ~)
function isolist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function Km_Callback(~, ~, ~)
function Km_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function AA_Callback(~, ~, ~)
function AA_CreateFcn(hObject, ~, ~)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function Vmax_Callback(~, ~, ~)
function Vmax_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function rVmax_Callback(~, ~, ~)
function rVmax_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function Vol_Callback(~, ~, ~)
function Vol_CreateFcn(hObject, ~, ~)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function dt_Callback(~, ~, ~)
function dt_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function popupmenu5_Callback(~, ~, ~)
function popupmenu5_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function exportfolder_Callback(~, ~, ~)
function exportfolder_CreateFcn(hObject, ~, ~)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function AAcon_Callback(~, ~, ~)
function AAcon_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function AAdep_Callback(~, ~, ~)
function AAdep_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function AATAAlist_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function transporters_CreateFcn(hObject, ~, ~)
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function EAAT4_Callback(~, ~, ~)
function EAAT5_Callback(~, ~, ~)
function SNAT5_Callback(~, ~, ~)
function GlyT1_Callback(~, ~, ~)
function ATB0_Callback(~, ~, ~)
function NTT4_Callback(~, ~, ~)
function XT2_Callback(~, ~, ~)
function B0AT2_Callback(~, ~, ~)
function B0AT1_Callback(~, ~, ~)
function CAT3_Callback(~, ~, ~)
function CAT4_Callback(~, ~, ~)
function b0AT_Callback(~, ~, ~)
function Asc1_Callback(~, ~, ~)
function Asc2_Callback(~, ~, ~)
function AGT1_Callback(~, ~, ~)
function TAT1_Callback(~, ~, ~)
function SNAT7_Callback(~, ~, ~)
function SNAT8_Callback(~, ~, ~)
function SIT1_Callback(~, ~, ~)
function step_Callback(~, ~, ~)
function PROT_Callback(hObject, eventdata, handles)
function step_CreateFcn(hObject, ~, ~)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function range_Callback(~, ~, ~)
function range_CreateFcn(hObject, ~, ~)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
