function[dAAindt]=ElectrogenicAntiporter(AAex,AAin,Kmex,Kmin,rVmax,Vmax,dt,Vol,B)

% Apparent Km
Kappin=zeros(1,22);
Kappex=zeros(1,22);
Kmin(isnan(Kmin))=inf;
Kmex(isnan(Kmex))=inf;
rVmax(isnan(rVmax))=0;
Xin=zeros(1,22);
Xex=zeros(1,22);

for i=1:22
    for n=1:22
        Xin(n)=(AAin(n)/Kmin(n));
        Xex(n)=(AAex(n)/Kmex(n));
    end
    Kappin(i)=Kmin(i)*(1+sum(Xin)-(AAin(i)/Kmin(i)));
    Kappex(i)=Kmex(i)*(1+sum(Xex)-(AAex(i)/Kmex(i)));
end

% Fractional Saturation
Fsin=zeros(1,22);
Fsex=zeros(1,22);

for i=1:22
    Fsin(i)=AAin(i)/(Kappin(i)+AAin(i));
    Fsex(i)=AAex(i)/(Kappex(i)+AAex(i));    
end

% Individual  Flux: AAex(n)|AAin(i)
%column=flux of this AA; row=due to this AA

Jio=zeros(22,22);
Joi=zeros(22,22);

for i=1:22
    for n=1:22
        %Condition: AA0ex|AA0in
        if  (i~=7 && i~=9 && i~=13 && i~=16)
            if (n~=7 && n~=9 && n~=13 && n~=16)
                Joi(i,n)=Vmax*Fsex(n)*Fsin(i)*rVmax(n)*dt;
                Jio(n,i)=Vmax*Fsin(i)*Fsex(n)*rVmax(i)*dt;
                %Condition: AA+ex|AA0in
            elseif (n==7 || n==9 || n==13 || n==16)
                Joi(i,n)=(B)*Vmax*Fsex(n)*Fsin(i)*rVmax(n)*dt;
                Jio(n,i)=(B)*Vmax*Fsin(i)*Fsex(n)*rVmax(i)*dt;
            end
            %Condition: AA0ex|AA+in
        elseif (i==7 || i==9 || i==13 || i==16)
            if (n~=7 && n~=9 && n~=13 && n~=16)
                Joi(i,n)=(1/B)*Vmax*Fsex(n)*Fsin(i)*rVmax(n)*dt;
                Jio(n,i)=(1/B)*Vmax*Fsin(i)*Fsex(n)*rVmax(i)*dt;
                %Condition: AA+ex|AA+in
            elseif   (n==7 || n==9 || n==13 || n==16)
                Joi(i,n)=Vmax*Fsex(n)*Fsin(i)*rVmax(n)*dt;
                Jio(n,i)=Vmax*Fsin(i)*Fsex(n)*rVmax(i)*dt;
            end
        end
        if (Joi(i,n)==0)||(Jio(n,i)==0)
            Joi(i,n)=0;
            Jio(n,i)=0;
        end
    end
end

% Calculate total flux for each AA (sum of all individual fluxes)
Jaaoi=sum(Joi);
Jaaio=sum(Jio);

% Change in AA concentration for each compartment
dAAindt=zeros(1,22);

for i=1:22
    dAAindt(i)= (Jaaoi(i)-Jaaio(i))/Vol;
end

end