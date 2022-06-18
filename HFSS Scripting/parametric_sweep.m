
% ----------------------------------------------
% Script Written by Ömer Katý for automation of creating randomly-generated sweeps
% 21:56:41  06.08.2020
% ----------------------------------------------
%% 
clear,clc;
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
    '\n'' Script Written by Ömer Katý for automation of creating randomly-generated sweeps'...
    '\n'' %s:%s:%s  %s.%s.%s'...
    '\n'' ----------------------------------------------'];
%%
Goals = '), _\nArray("NAME:Goals"))';
%%
AnsoftApp.Matrix = input('Cell Array (Ex: matlab.mat) = ', 's');
data = struct2cell(load(AnsoftApp.Matrix));
data = data{1,1};
%%
AnsoftApp.Project = input('Active Project (Ex: Project1) = ', 's');
AnsoftApp.Design = input('Active Design (Ex: HFSSDesign1) = ', 's');
AnsoftApp.ParametricSweep = input('Parametric Sweep (Ex: ParametricSetup1) = ', 's');
AnsoftApp.IsEnabled = input('Setup Status (Ex: true/false) = ', 's');
AnsoftApp.SaveFields = input('Save Fields And Mesh (Ex: true/false) = ', 's');
AnsoftApp.Sim_Setups = input('Solution Setup (Ex: ATK_Solution) = ', 's');
AnsoftApp.ScriptFile = input('Script File (Ex: Script1.txt) = ', 's');
%%
fileID = fopen(AnsoftApp.ScriptFile,'w');
%%
dataInfo = size(data);
AnsoftApp.Inputs.Size = dataInfo(2);
AnsoftApp.Data.Size = dataInfo(1)-2;
%%
AnsoftApp.Parameters = strings([1 AnsoftApp.Inputs.Size]);
AnsoftApp.Units = strings([1 AnsoftApp.Inputs.Size]);
Sweeps = "";
%%
for i=1:1:AnsoftApp.Inputs.Size
    if(data{1,i}() ~= "Freq" && data{1,i}() ~= "Phase")
        AnsoftApp.Parameters(1,i) = string(data{1,i}());
        AnsoftApp.Units(1,i) = string(data{2,i}());
        Sweeps = strcat(Sweeps,strjoin([', _\nArray("NAME:SweepDefinition", "Variable:=", "'...
            AnsoftApp.Parameters(1,i)...
            '", "Data:=", "0'...
            AnsoftApp.Units(1,i)...
            '", "OffsetF1:=", false, "Synchronize:=", 0)'],''));
    end

end
%%
startFormat = ['\nDim oAnsoftApp'...
    '\nDim oDesktop'...
    '\nDim oProject'... 
    '\nDim oDesign'... 
    '\nDim oEditor'... 
    '\nDim oModule'... 
    '\nSet oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")'... 
    '\nSet oDesktop = oAnsoftApp.GetAppDesktop()'... 
    '\noDesktop.RestoreWindow'... 
    '\nSet oProject = oDesktop.SetActiveProject("' AnsoftApp.Project    '")'... 
    '\nSet oDesign = oProject.SetActiveDesign("' AnsoftApp.Design '")'... 
    '\nSet oModule = oDesign.GetModule("Optimetrics")'...
    '\noModule.InsertSetup "OptiParametric", _'...
    '\nArray("NAME:' AnsoftApp.ParametricSweep '", "IsEnabled:=", '...
    AnsoftApp.IsEnabled ', _'...
    '\nArray("NAME:ProdOptiSetupDataV2", "SaveFields:=", ' AnsoftApp.SaveFields...
    ', "CopyMesh:=", false, "SolveWithCopiedMeshOnly:=", true), _'...
    '\nArray("NAME:StartingPoint"), "Sim. Setups:=", _'...
    '\nArray("' AnsoftApp.Sim_Setups '"), _'...
    '\nArray("NAME:Sweeps"'...
    Sweeps...
    '), _\nArray("NAME:Sweep Operations"'];
startFormat = strjoin(startFormat,'');
%%
AnsoftApp.bytes = fprintf(fileID,...
    Header,time(1,4),time(1,5),time(1,6),time(1,3),time(1,2),time(1,1));
AnsoftApp.bytes = AnsoftApp.bytes + fprintf(fileID,startFormat);
%%
for j=1:1:AnsoftApp.Data.Size
    AnsoftApp.bytes = AnsoftApp.bytes + ...
        fprintf(fileID,', _\n"add:=", Array(');
    for i=1:1:AnsoftApp.Inputs.Size
        if(data{1,i}() ~= "Freq" && data{1,i}() ~= "Phase")
            if(i == AnsoftApp.Inputs.Size)
                AnsoftApp.bytes = AnsoftApp.bytes + ...
                    fprintf(fileID,' "%f%s")',data{j+2,i}(),AnsoftApp.Units(1,i));
                break;
            end
            AnsoftApp.bytes = AnsoftApp.bytes + ...
                fprintf(fileID,' "%f%s",',data{j+2,i}(),AnsoftApp.Units(1,i));
        end
    end
end
%%
AnsoftApp.bytes = AnsoftApp.bytes + fprintf(fileID,Goals);
fclose(fileID);
%%
