function varargout = AutoSecond_1_5(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoSecond_1_5_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoSecond_1_5_OutputFcn, ...
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


% --- Executes just before AutoSecond_1_5 is made visible.
function AutoSecond_1_5_OpeningFcn(hObject, eventdata, handles, varargin)

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

function varargout = AutoSecond_1_5_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;






% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global str1
[filename,pathname,fileindex]=uigetfile('*.doc;*.docx','选择报告');
str1=fullfile(pathname,filename);
set(handles.edit1,'string',str1)

% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global str2
[filename,pathname,fileindex]=uigetfile('*.doc;*.docx','选择报告');
str2=fullfile(pathname,filename);
set(handles.edit2,'string',str2)

% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global str3
[filename,pathname,fileindex]=uigetfile('*.doc;*.docx','选择报告');
str3=fullfile(pathname,filename);
set(handles.edit3,'string',str3)


% --- Executes on button press in pushbutton4.
function pushbutton4_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global str4
[filename,pathname,fileindex]=uigetfile('*.doc;*.docx','选择报告');
str4=fullfile(pathname,filename);
set(handles.edit4,'string',str4)

% --- Executes on button press in pushbutton5.
function pushbutton5_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
t1=waitbar(0,'正在合成报告');
global str1 str2 str3 str4
if isempty(str1)||isempty(str2)||isempty(str3)||isempty(str4)
    msgbox('缺少某个报告');
    return;
end
file_usr=str1;
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=[cd,'\report.docx'];
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
waitbar(0.1);
%%%粘贴报告2%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible = 0; % 使word为可见；或set(Word, 'Visible', 1); 
Document=Word.Documents.Open(str2);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
 Selection.WholeStory
Selection.Copy
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Word.Quit; % 关闭文档

 waitbar(0.3);
Word=actxserver('Word.Application'); 

Word.Visible = 0; % 使word为可见；或set(Word, 'Visible', 1); 
Document=Word.Documents.Open(filespec_user);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
 Selection.WholeStory
Selection.EndKey
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Paste

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档

waitbar(0.4);
%%%粘贴报告3%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5

Word=actxserver('Word.Application'); 

Word.Visible = 0; % 使word为可见；或set(Word, 'Visible', 1); 
Document=Word.Documents.Open(str3);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
 Selection.WholeStory
Selection.Copy
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Word.Quit; % 关闭文档

waitbar(0.5);
Word=actxserver('Word.Application'); 

Word.Visible = 0; % 使word为可见；或set(Word, 'Visible', 1); 
Document=Word.Documents.Open(filespec_user);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
 Selection.WholeStory
Selection.EndKey
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Paste

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
waitbar(0.6);

%%%粘贴报告4%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5

Word=actxserver('Word.Application'); 

Word.Visible = 0; % 使word为可见；或set(Word, 'Visible', 1); 
Document=Word.Documents.Open(str4);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
 Selection.WholeStory
Selection.Copy
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Word.Quit; % 关闭文档
waitbar(0.8);
 
Word=actxserver('Word.Application'); 

Word.Visible = 0; % 使word为可见；或set(Word, 'Visible', 1); 
Document=Word.Documents.Open(filespec_user);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
 Selection.WholeStory
Selection.EndKey
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Paste

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
waitbar(0.9);
winopen(filespec_user);
waitbar(1)
close(t1);


function edit1_Callback(hObject, eventdata, handles)
% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function edit2_Callback(hObject, eventdata, handles)
% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

function edit3_Callback(hObject, eventdata, handles)
% --- Executes during object creation, after setting all properties.
function edit3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)

function edit4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
