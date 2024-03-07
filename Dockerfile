FROM dockurr/windows:2.04

COPY ./win11x64.xml /run/assets/
#COPY ./iso/tiny11_core_x64_beta_1.iso /storage/custom.iso
#COPY ./iso/SW_DVD5_Office_Professional_Plus_2010w_SP1_64Bit_English_CORE_MLF_X17-76756.ISO /storage/office.iso
COPY ./storage/shared /storage/shared/

ENV MANUAL: "N"
ENV CPU_CORES "2"
ENV DISK_SIZE "16G"
ENV VERSION "win11"
# assumes disk is mounted on D:, driver mounted on E: during installation
ENV ARGUMENTS "-drive id=cdrom3,media=cdrom,if=none,format=raw,readonly=on,file=/storage/office.iso,index=9 -device ide-cd,drive=cdrom3,bus=ide.3,unit=0"