
% ----------------------------------------------
% Script Written by �mer Kati for automating Fields Calculator use
% 21:56:41  06.08.2020
% ----------------------------------------------
%%
% clear,clc;
%% 
time = fix(clock);
for i=2:6
    if(time(1,i)<10) time(2,i) = 1;, end
end
%% 
time = string(time);
%% 
for i=2:6
    if(time(2,i) == "1") time(1,i) = strcat('0',time(1,i));, end
end
%%
Header = ['\n'' ----------------------------------------------'...
    '\n'' Script Written by Omer Kati for automating Fields Calculator use'...
    '\n'' %s:%s:%s  %s.%s.%s'...
    '\n'' ----------------------------------------------']
%%
AnsoftApp.Matrix = input('Cell Array (Ex: matlab.mat) = ', 's');
data = struct2cell(load(AnsoftApp.Matrix));
data = data{1,1};
%%
AnsoftApp.Project = input('Active Project (Ex: Project1) = ', 's');
AnsoftApp.Design = input('Active Design (Ex: HFSSDesign1) = ', 's');
AnsoftApp.Module = input('Active Module (Ex: FieldsReporter) = ', 's');
AnsoftApp.Solution = input('Solution (Ex: ATK_Solution : FF_Sweep) = ', 's');
AnsoftApp.Surface = input('Surface (Ex: Rectangle1) = ', 's');
AnsoftApp.ScriptFile = input('Script File (Ex: Script1.txt) = ', 's');
%%
AnsoftApp.Inputs.Size = size(data,2); 
AnsoftApp.Data.Size = size(data,1)-2;
%%
fileID = fopen(AnsoftApp.ScriptFile,'w');
%%
for i=1:AnsoftApp.Inputs.Size
    
    AnsoftApp.Parameters(1,i) = string(data{1,i}());
    AnsoftApp.Units(1,i) = string(data{2,i}());

end
%%

Declarations = string(['\nDim oAnsoftApp'...
    '\nDim oDesktop'...
    '\nDim oProject'... 
    '\nDim oDesign'... 
    '\nDim oEditor'... 
    '\nDim oModule'... 
    '\nSet oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")'... 
    '\nSet oDesktop = oAnsoftApp.GetAppDesktop()'... 
    '\noDesktop.RestoreWindow'... 
    '\nSet oProject = oDesktop.SetActiveProject("' AnsoftApp.Project '")'... 
    '\nSet oDesign = oProject.SetActiveDesign("' AnsoftApp.Design '")'... 
    '\nSet oModule = oDesign.GetModule("' AnsoftApp.Module '")'...
    '\noModule.EnterQty "Poynting"'...
    '\noModule.CalcOp "Mag"'...
    '\noModule.EnterSurf "' AnsoftApp.Surface '"'...
    '\noModule.CalcOp "Integrate"'...
    '\noModule.EnterSurf "' AnsoftApp.Surface '"'...
    '\noModule.EnterScalar 1'...
    '\noModule.CalcStack "exch"'...
    '\noModule.CalcOp "Integrate"'...
    '\noModule.CalcOp "/"']);
%%
AnsoftApp.bytes = fprintf(fileID,...
    Header,time(1,4),time(1,5),time(1,6),time(1,3),time(1,2),time(1,1));
AnsoftApp.bytes = AnsoftApp.bytes + fprintf(fileID,Declarations);
%%
for j = 1:AnsoftApp.Data.Size
    AnsoftApp.bytes = AnsoftApp.bytes ...
        + fprintf(fileID,'\noModule.ClcEval "%s", Array( _',...
        AnsoftApp.Solution);
    for i=1:1:AnsoftApp.Inputs.Size
        if(i == AnsoftApp.Inputs.Size)
            AnsoftApp.bytes = AnsoftApp.bytes...
                + fprintf(fileID,...
                '\n"%s:=", "%f%s")\noModule.CalcStack "exch"',...
                AnsoftApp.Parameters(1,i),data{j+2,i}(),...
                AnsoftApp.Units(1,i));
            break;
        end
        AnsoftApp.bytes = AnsoftApp.bytes...
            + fprintf(fileID,'\n"%s:=", "%f%s", _',...
            AnsoftApp.Parameters(1,i),data{j+2,i}(),...
            AnsoftApp.Units(1,i));
    end
    
end
%%
fclose(fileID);
%%
